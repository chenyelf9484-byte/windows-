import re
from dataclasses import dataclass


@dataclass(frozen=True)
class Rule:
    source: str
    target: str


class Matcher:
    def __init__(self, rules, ignore_case=False):
        self.rules = tuple(rules)
        if not self.rules or any(not r.source for r in self.rules):
            raise ValueError('至少设置一条非空关键词规则')
        self.ignore_case = ignore_case
        keys = [r.source.casefold() if ignore_case else r.source for r in self.rules]
        if len(set(keys)) != len(keys):
            raise ValueError('存在重复关键词，请删除冲突规则')
        # 每个分支单独命名，避免 Unicode 大小写匹配后的字典查找歧义。
        ordered = sorted(enumerate(self.rules), key=lambda x: -len(x[1].source))
        pattern = '|'.join('(?P<r%d>%s)' % (i, re.escape(r.source)) for i, r in ordered)
        self.pattern = re.compile(pattern, re.IGNORECASE if ignore_case else 0)

    def matches(self, text):
        return [(m.start(), m.end(), self.rules[int(m.lastgroup[1:])].target)
                for m in self.pattern.finditer(text)]

    def replace(self, text):
        return self.pattern.sub(lambda m: self.rules[int(m.lastgroup[1:])].target, text)

    def to_json(self):
        return {'ignore_case': self.ignore_case,
                'rules': [{'source': r.source, 'target': r.target} for r in self.rules]}
