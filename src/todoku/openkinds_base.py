"""openkinds と morekinds が共に使う例外（入れ子の読み込みを避けるための小さな部品）。"""


class OpenFormatError(ValueError):
    pass
