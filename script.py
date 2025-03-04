import pandas as pd
import argparse
import re
from ruamel.yaml import YAML
from ruamel.yaml.scalarstring import DoubleQuotedScalarString as dq

def extract_namespace(excel_path):
    """Извлекает namespace из ячейки A2 Excel-файла."""
    df_meta = pd.read_excel(excel_path, sheet_name='Лист1', header=None, nrows=2)
    namespace_line = df_meta.iloc[1, 0]
    if not isinstance(namespace_line, str) or "namespace:" not in namespace_line:
        raise ValueError("Ошибка: в A2 должен быть указан namespace в формате 'namespace: <значение>'")
    return namespace_line.split(":")[1].strip()

def is_fqdn(source: str) -> bool:
    """Проверяет, является ли значение FQDN (например, ya.ru)."""
    return '.' in source and '/' not in source and not source.startswith('app:')

def is_cidr(source: str) -> bool:
    """Проверяет, является ли значение CIDR (например, 10.30.24.1/32)."""
    cidr_pattern = r"^(\d{1,3}\.){3}\d{1,3}/\d{1,2}$|^([0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}/\d{1,3}$"
    return re.match(cidr_pattern, source) is not None

def parse_namespace_pod(value: str) -> dict:
    """Обрабатывает значения вида 'namespace/pod' (например, kube-system/kube-dns)."""
    if '/' in value and not is_cidr(value):
        namespace, pod_label = value.split('/', 1)
        return {'namespace': namespace, 'pod_label': pod_label}
    return None

def main():
    # Настройка парсера аргументов командной строки
    parser = argparse.ArgumentParser(description='Генерация политик Cilium из Excel.')
    parser.add_argument('file', type=str, help='Путь к Excel-файлу.')
    args = parser.parse_args()

    fqdn_ingress_errors = []  # Список для ошибок FQDN в ingress
    invalid_endpoints = []    # Список для некорректных endpoint

    try:
        # Извлечение namespace из Excel
        namespace = extract_namespace(args.file)
        output_filename = f"{namespace}.yaml"
    except Exception as e:
        print(f"Ошибка при чтении namespace: {e}")
        return

    try:
        # Чтение данных из Excel
        df = pd.read_excel(args.file, sheet_name='Лист1')
        expected_columns = [
            '№пп', 
            'Протокол (TCP,UDP)', 
            'endpoint', 
            'ingress/egress', 
            'Сервис', 
            'Порт', 
            'Описание правила'
        ]
        if not all(col in df.columns for col in expected_columns):
            raise ValueError(f"Неверный формат колонок. Ожидаемые колонки: {', '.join(expected_columns)}")
        df = df[pd.to_numeric(df['№пп'], errors='coerce').notna()]
    except Exception as e:
        print(f"Ошибка при чтении данных: {e}")
        return

    # Группировка данных по endpoint
    grouped = df.groupby('endpoint')
    policies = []
    yaml = YAML()
    yaml.indent(sequence=4, offset=2)  # Настройка форматирования YAML

    for endpoint, group in grouped:
        try:
            # Пропуск endpoint без метки (например, "app:test1")
            if ':' not in endpoint:
                invalid_endpoints.append(endpoint)
                print(f"Пропущен некорректный endpoint: {endpoint}")
                continue

            # Формирование имени политики
            label_key, label_value = endpoint.split(':', 1)
            policy_name = f"{namespace}-{label_value.strip()}"

            # Базовая структура политики
            policy = {
                'apiVersion': 'cilium.io/v2',
                'kind': 'CiliumNetworkPolicy',
                'metadata': {
                    'name': policy_name,
                    'namespace': namespace
                },
                'spec': {
                    'endpointSelector': {'matchLabels': {label_key: label_value}},
                    'ingress': [],
                    'egress': []
                }
            }

            # Обработка каждой строки в группе
            for _, row in group.iterrows():
                try:
                    rule_number = int(row['№пп'])
                    direction = row['ingress/egress'].strip().lower()
                    protocol = row['Протокол (TCP,UDP)'].strip().upper()
                    port = int(row['Порт'])
                    source_dest = row['Сервис'].strip()

                    # Формирование правила для порта
                    port_rule = {
                        'port': dq(port),
                        'protocol': dq(protocol)
                    }

                    # Обработка значения "Сервис"
                    ns_pod = parse_namespace_pod(source_dest)
                    rule = {}

                    # Правила для ingress
                    if direction == 'ingress':
                        if is_fqdn(source_dest):
                            fqdn_ingress_errors.append(rule_number)
                            print(f"FQDN в ingress игнорируется (правило {rule_number})")
                            continue

                        if ns_pod:
                            rule['fromEndpoints'] = [{
                                'matchLabels': {
                                    'namespace': ns_pod['namespace'],
                                    'app': ns_pod['pod_label']
                                }
                            }]
                        elif is_cidr(source_dest):
                            rule['fromCIDRSet'] = [{'cidr': source_dest}]
                        elif source_dest.startswith('app:'):
                            src_key, src_value = source_dest.split(':', 1)
                            rule['fromEndpoints'] = [{'matchLabels': {src_key: src_value}}]
                        else:
                            print(f"Неподдерживаемый источник ingress: {source_dest}")
                            continue

                        rule['toPorts'] = [{'ports': [port_rule]}]
                        policy['spec']['ingress'].append(rule)

                    # Правила для egress
                    elif direction == 'egress':
                        if ns_pod:
                            rule['toEndpoints'] = [{
                                'matchLabels': {
                                    'namespace': ns_pod['namespace'],
                                    'app': ns_pod['pod_label']
                                }
                            }]
                        elif is_cidr(source_dest):
                            rule['toCIDRSet'] = [{'cidr': source_dest}]
                        elif source_dest.startswith('app:'):
                            dst_key, dst_value = source_dest.split(':', 1)
                            rule['toEndpoints'] = [{'matchLabels': {dst_key: dst_value}}]
                        elif is_fqdn(source_dest):
                            rule['toFQDNs'] = [{'matchName': source_dest}]
                        else:
                            print(f"Неподдерживаемый приемник egress: {source_dest}")
                            continue

                        rule['toPorts'] = [{'ports': [port_rule]}]
                        policy['spec']['egress'].append(rule)

                except ValueError:
                    print(f"Ошибка в строке {row['№пп']}: порт должен быть целым числом.")
                except Exception as e:
                    print(f"Ошибка в строке {row['№пп']}: {e}")

            # Добавление пустых правил, если списки пусты
            if not policy['spec']['ingress']:
                policy['spec']['ingress'] = [{}]  # ingress: [ - {} ]
            if not policy['spec']['egress']:
                policy['spec']['egress'] = [{}]   # egress: [ - {} ]

            policies.append(policy)

        except Exception as e:
            print(f"Ошибка при обработке endpoint {endpoint}: {e}")

    # Запись политик в YAML-файл
    with open(output_filename, 'w', encoding='utf-8') as f:
        # Запись ошибок в начало файла
        error_msgs = []
        if invalid_endpoints:
            error_msgs.append(f"# Пропущены некорректные endpoint: {', '.join(invalid_endpoints)}")
        if fqdn_ingress_errors:
            error_msgs.append(f"# FQDN в ingress игнорируются (номера правил): {', '.join(map(str, sorted(fqdn_ingress_errors)))}")
        
        if error_msgs:
            f.write("\n".join(error_msgs) + "\n---\n")
        
        yaml.dump_all(policies, f)

    print(f"Файл создан: {output_filename}")

if __name__ == '__main__':
    main()