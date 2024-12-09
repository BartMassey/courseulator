import re, sys

vowel_re = re.compile(r"[AEIOU]")
split_re = re.compile(r"[^A-Z]")

def tileset(s):
    result = set()
    for word in split_re.split(s):
        w = vowel_re.sub("", word)
        if len(w) < 3:
            w = word
        nw = len(w)
        if nw < 3:
            continue
        tiles = set(w[i:i+3] for i in range(nw - 2))
        result |= tiles
    return result

def sim(ts1, ts2):
    if not ts1 and not ts2:
        return 0
    return len(ts1 & ts2) / len(ts1 | ts2)

if __name__ == "__main__":
    ts1 = tileset("RUST EMBEDDED")
    print(ts1)
    ts2 = tileset("RUST EMBEDDEX")
    print(ts2)
    print(sim(ts1, ts2))
