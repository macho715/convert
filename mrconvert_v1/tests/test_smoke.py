from mrconvert.cli import build_parser

def test_parser_builds():
    p = build_parser()
    assert p is not None
