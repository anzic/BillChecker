# -*- coding: UTF-8 -*-
import json

class ClassifyRule():
    def __init__(self, rule_dict):
        'init rule with dict'
        self.attr = rule_dict['attr']
        self.relation = rule_dict['relation']
        self.value = rule_dict['value']
        self.type = rule_dict['type']


def JsonParser(fname):
    rules = []
    f = open(fname, 'r', encoding='utf-8')
    text = f.read()
    f.close()

    rule_dicts = json.loads(text)['rules']
    for rule_dict in rule_dicts:
        rules.append(ClassifyRule(rule_dict))
    return rules


if __name__ == "__main__":
    rules = JsonParser('classify_rule.json')
    pass

