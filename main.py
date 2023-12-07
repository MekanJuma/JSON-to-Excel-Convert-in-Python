import pandas as pd
import json
import os

class JSONToExcelConverter:

    @staticmethod
    def parse_rules(rules):
        if 'condition' in rules:
            return f"{rules['field']['label']} = {rules['data']}"
        elif 'group' in rules:
            group_rules = ' {operator} '.join([JSONToExcelConverter.parse_rules(r) for r in rules['group']['rules']])
            return f"({group_rules.format(operator=rules['group']['operator'].replace('&&', 'AND').replace('||', 'OR'))})"
        elif 'rules' in rules:
            rule_operator = ' {operator} '.format(operator=rules['operator'].replace('&&', 'AND').replace('||', 'OR'))
            return rule_operator.join([JSONToExcelConverter.parse_rules(r) for r in rules['rules']])
        return ''

    def transform_to_dataframe(self, json_data, filename):
        data = []
        for page in json_data.get('pages', []):
            page_name = page.get('name', '')
            page_filter = JSONToExcelConverter.parse_rules(page.get('rule', {}))
            for row in page.get('rows', []):
                section = row.get('label', '')
                row_filter = JSONToExcelConverter.parse_rules(row.get('filterrule', {}))
                for item in row.get('items', []):
                    verbiage = item.get('label', '')
                    content = item.get('content', '')
                    field = item.get('name', '')
                    obj = item.get('object', '')
                    response_type = item.get('type', item.get('componenttype'))
                    required = item.get('required', '')
                    options = ', '.join([option['label'] for option in item.get('options', [])[:10]]) + ('...' if len(item.get('options', [])) > 10 else '')
                    selectedOptions = ', '.join([option['label'] for option in item.get('selectOptions', [])[:10]]) + ('...' if len(item.get('selectOptions', [])) > 10 else '')
                    field_filter = JSONToExcelConverter.parse_rules(item.get('filterrule', item.get('rule', {})))
                    all_filters = field_filter + row_filter + page_filter

                    applies_trad = 'YES' if 'Delivery Method = On Campus - Trad' in all_filters else 'NO'
                    applies_online = 'YES' if 'Delivery Method = Online' in all_filters else 'NO'
                    applies_saoe = 'YES' if 'Delivery Method = On Campus - SAOE' in all_filters else 'NO'
                    if 'paramedic' not in page_name.lower():
                        data.append([filename.split('/')[-1].split('.')[0], page_name, section, obj, field, verbiage, content, response_type, options, selectedOptions, required, field_filter, row_filter, page_filter, applies_trad, applies_online, applies_saoe])

        columns = ['Form name', 'Page', 'Section', 'Object', 'Field', 'Verbiage', 'Content', 'Response Type', 'Available Options', 'Selected Options', 'Required?', 'Field Filter', 'Row Filter', 'Page Filter', 'Applies to Trad?', 'Applies to Online?', 'Applies to SAOE?']
        return pd.DataFrame(data, columns=columns)

    def read_and_process_all(self, save_to_excel=False):
        all_data = []
        json_files = [f for f in os.listdir('files') if f.endswith('.json')]
        for json_file in json_files:
            filename = os.path.join('files', json_file)
            with open(filename, 'r') as file:
                json_data = json.load(file)
            df = self.transform_to_dataframe(json_data, filename)
            all_data.append(df)
        combined_df = pd.concat(all_data, ignore_index=True)
        combined_df = combined_df.sort_values(by='Form name', ascending=True)
        if save_to_excel:
            combined_df.to_excel('form analysis.xlsx', index=False)
        else:
            print(combined_df)


converter = JSONToExcelConverter()
converter.read_and_process_all(save_to_excel=True)
