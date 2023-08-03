import pandas as pd


def find_overlap_fragments(content0: str, content1: str, min_fragment_size):
    if pd.isna(content0) or pd.isna(content1):
        return set()
    shorter, longer = (content0, content1) if len(content0) < len(content1) else (content1, content0)
    fragments = set()
    for i in range(len(shorter)):
        start = i + min_fragment_size
        end = len(shorter) + 1
        while start < end:
            mid = (start + end) // 2
            fragment = shorter[i:mid]
            if fragment in longer:
                fragments.add(fragment)
                start = mid + 1
            else:
                end = mid
    return list(filter(lambda x: all(x not in y for y in fragments if x != y), fragments))
