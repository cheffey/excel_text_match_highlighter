import json
import re


class ExcelRichText:
    def __init__(self, text: str):
        self.text = str(text)
        self.marks = []

    def color(self, part, part_color, only_color_blank=False):
        self.is_colored(part, part_color, only_color_blank=only_color_blank)
        return self

    def is_colored(self, target, part_color, only_color_blank=False):
        colored = False
        if not isinstance(target, str):
            for tar in target:
                if self.is_colored(tar, part_color, only_color_blank=only_color_blank):
                    colored = True
            return colored
        level = len(self.marks)
        is_regex = target.startswith('(') and target.endswith(')')
        if is_regex:
            parts = re.findall(target, self.text)
        else:
            parts = [target]
        for part in parts:
            skip = 0
            while True:
                start = self.text.find(part, skip)
                if start == -1:
                    break
                end = start + len(part)
                mark = Mark(start, end, part_color, level, only_color_blank, self.text[start:end])
                self.marks.append(mark)
                colored = True
                skip = end
        return colored

    def write_to(self, sheet, row, col):
        content, params = self._build_text_param()
        if len(params) == 1:
            sheet.write(row, col, params[0])
        elif len(params) == 2:
            sheet.write(row, col, params[1], params[0])
        else:
            sheet.write_rich_string(row, col, *params)

    def _get_one_overlap_mark_pair(self):
        for i, mark in enumerate(self.marks):
            others = self.marks[i + 1:]
            for other in others:
                if mark.start < other.start < mark.end:
                    return mark, other
                elif other.start < mark.start < other.end:
                    return other, mark
                elif mark.start == other.start:
                    if mark.end > other.end:
                        return mark, other
                    else:
                        return other, mark
        return None, None

    def _reformat_marks(self):
        while True:
            a1, a2 = self._get_one_overlap_mark_pair()
            if a1 is None:
                break
            if a1.level == a2.level or a1.color == a2.color:
                a1.end = a2.end
                self.marks.remove(a2)
            elif a1.start == a2.start and a1.end == a2.end:
                if a1.level > a2.level:
                    self.marks.remove(a1)
                else:
                    self.marks.remove(a2)
            elif a2.end <= a1.end:
                if a2.only_color_blank:
                    self.marks.remove(a2)
                else:
                    self.marks.remove(a1)
                    if a1.start < a2.start:
                        self.marks.append(Mark(a1.start, a2.start, a1.color, a1.level, a1.only_color_blank, self.text[a1.start:a2.start]))
                    if a2.end < a1.end:
                        self.marks.append(Mark(a2.end, a1.end, a1.color, a1.level, a1.only_color_blank, self.text[a2.end:a1.end]))
            elif a1.level < a2.level:
                a2.start = a1.end
            else:
                a1.end = a2.start

    def _build_text_param(self):
        self._reformat_marks()
        marks = list(self.marks)
        marks.sort(key=lambda x: x.start)
        content = []
        since = 0
        for mark in marks:
            content.append((None, self.text[since:mark.start]))
            content.append((mark.color, self.text[mark.start:mark.end]))
            since = mark.end
        content.append((None, self.text[since:]))
        params = []
        content = [c for c in content if c[1]]
        for color, text in content:
            if text and color is None:
                params.append(text)
            elif text:
                params.append(color)
                params.append(text)
        return content, params


class Mark:
    def __init__(self, start, end, color, level, only_color_blank, text):
        self.start = start
        self.end = end
        self.color = color
        self.level = level
        self.only_color_blank = only_color_blank
        self.text = text

    def __str__(self):
        return f'{self.start}-{self.end} {self.text}'


if __name__ == '__main__':
    bing_json = json.loads("""
    {"score": 0, "feedback": "\u63d0\u4ea4\u7684\u7b54\u6848\u4e0e\u4e13\u5bb6\u7b54\u6848\u4e0d\u7b26\uff0c\u4e14\u6709\u591a\u5904\u4e8b\u5b9e\u9519\u8bef", "pros": [], "cons": ["\u516c\u5143\u524d260\u5e74\u81f3\u524d237\u5e74", "\u79e6\u3001\u8d75\u3001\u9b4f\u4e09\u4e2a\u8bf8\u4faf\u56fd\u74dc\u5206\u7684\u4e8b\u4ef6", "\u79e6\u56fd\u593a\u53d6\u4e86\u664b\u56fd\u5317\u90e8\u5730\u533a\uff0c\u5305\u62ec\u90fd\u57ce\u4e34\u6c7e\uff1b\u8d75\u56fd\u593a\u53d6\u4e86\u664b\u56fd\u4e2d\u90e8\u5730\u533a\uff1b\u9b4f\u56fd\u593a\u53d6\u4e86\u664b\u56fd\u5357\u90e8\u5730\u533a"], "unrealities": ["\u516c\u5143\u524d260\u5e74\u81f3\u524d237\u5e74", "\u79e6\u3001\u8d75\u3001\u9b4f\u4e09\u4e2a\u8bf8\u4faf\u56fd\u74dc\u5206\u7684\u4e8b\u4ef6", "\u79e6\u56fd\u593a\u53d6\u4e86\u664b\u56fd\u5317\u90e8\u5730\u533a\uff0c\u5305\u62ec\u90fd\u57ce\u4e34\u6c7e\uff1b\u8d75\u56fd\u593a\u53d6\u4e86\u664b\u56fd\u4e2d\u90e8\u5730\u533a\uff1b\u9b4f\u56fd\u593a\u53d6\u4e86\u664b\u56fd\u5357\u90e8\u5730\u533a"], "keys": []}
    """)
    actual_answer = '三家分晋是指公元前260年至前237年，中国战国末年晋国被秦、赵、魏三个诸侯国瓜分的事件。晋国原本是春秋时期的一个强大诸侯国，但在战国时期国势衰落，最终被三个强国瓜分。秦国夺取了晋国北部地区，包括都城临汾；赵国夺取了晋国中部地区；魏国夺取了晋国南部地区。这一事件标志着晋国的灭亡，并为秦国统一六国奠定了基础。'
    score = bing_json['score']
    feedback = bing_json['feedback']
    pros = list(filter(lambda x: x, bing_json.get('pros', [])))
    keys = list(filter(lambda x: x, bing_json.get('keys', [])))
    cons = list(filter(lambda x: x, bing_json.get('cons', [])))
    unrealities = list(filter(lambda x: x, bing_json.get('unrealities', [])))
    actual_rich_text = ExcelRichText(actual_answer)
    not_marked_pros = []
    not_marked_cons = []
    for pro in pros:
        if not actual_rich_text.is_colored(pro, ['blue']):
            not_marked_pros.append(pro)
    for key in keys:
        if not actual_rich_text.is_colored(key, ['green']):
            not_marked_pros.append(key)
    for con in cons:
        if not actual_rich_text.is_colored(con, ['red']):
            not_marked_cons.append(con)
    for unreal in unrealities:
        if not actual_rich_text.is_colored(unreal, ['purple']):
            not_marked_cons.append(unreal)
    if len(not_marked_pros) > 0:
        feedback += f'\n\n正确: {not_marked_pros}'
    if len(not_marked_cons) > 0:
        feedback += f'\n\n错误: {not_marked_cons}'
    feedback_rich_text = ExcelRichText(feedback)
    [feedback_rich_text.color(text, ['blue']) for text in not_marked_pros]
    [feedback_rich_text.color(text, ['red']) for text in not_marked_cons]
    actual_rich_text._build_text_param()
