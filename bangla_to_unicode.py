import json
import logging
import re

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class BanglaToUnicode:
    def __init__(self):
        self.sutonnymj_mapper = None
        self.sutonnymj_path = "resources/sutonnymj_mapper.json"

    def unicode_to_ncr(self, text):
        # NCR = Numeric Character Reference Example: 'A' -> '&#65;'
        return "".join(f"&#{ord(char)};" for char in text)

    def ncr_to_unicode(self, text):
        # Example: '&#65;' -> 'A'
        return re.sub(r"&#(\d+);", lambda match: chr(int(match.group(1))), text)

    def to_sutonnymj(self, text):
        if not self.sutonnymj_mapper:
            with open(self.sutonnymj_path, "r") as file:
                self.sutonnymj_mapper = json.load(file)

        return self.convert(
            text, self.sutonnymj_mapper["find"], self.sutonnymj_mapper["replace"]
        )

    def from_sutonnymj(self, text):
        if not self.sutonnymj_mapper:
            with open(self.sutonnymj_path, "r") as file:
                self.sutonnymj_mapper = json.load(file)

        return self.convert(
            text, self.sutonnymj_mapper["replace"], self.sutonnymj_mapper["find"]
        )

    def convert(self, text, find_patterns, replace_patterns):
        if len(find_patterns) != len(replace_patterns):
            raise ValueError(
                "find_patterns and replace_patterns must have the same length."
            )
        text = self.unicode_to_ncr(text)
        for find, replace in zip(find_patterns, replace_patterns):
            text = re.sub(find, replace, text)

        return self.ncr_to_unicode(text)


if __name__ == "__main__":
    btu = BanglaToUnicode()
    text = "Bmjv‡gi mykxZj QvqvZ‡j Avkªq MÖnYKvix fviZxq bIgymwjg‡`i †ivgvÂKi mv¶vrKvi"
    expected = (
        "ইসলামের সুশীতল ছায়াতলে আশ্রয় গ্রহণকারী ভারতীয় নওমুসলিমদের রোমাঞ্চকর সাক্ষাৎকার"
    )
    converted = btu.to_sutonnymj(text)
    assert converted == expected, f"Expected: {expected},\n Got: {converted}"
    logger.info("Hurrah!! sutonny to unicode passed!")
    converted_back = btu.from_sutonnymj(converted)
    unicode_again = btu.to_sutonnymj(converted_back)
    assert unicode_again == expected, f"Expected: {expected},\n Got: {unicode_again}"
    logger.info("Hurrah!! unicode to sutonny passed!")
