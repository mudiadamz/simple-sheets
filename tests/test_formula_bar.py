"""The formula bar grows for long formulas and colours references like the grid outlines."""
from grid import REF_COLORS


def test_grows_and_colours(app):
    v = app.value
    app.update()
    assert int(v.cget("height")) == 1
    long = "=" + "+".join(f"A{i}" for i in range(1, 120))
    app.set_cells([(0, 5, long)])
    app.table.goto(0, 5)
    app.update()
    assert v.get() == long
    v.fit()
    assert 1 < int(v.cget("height")) <= v.MAX_LINES                    # several lines, but not unlimited
    app.table.goto(1, 5)
    app.update()
    v.fit()
    assert int(v.cget("height")) == 1
    app.set_cells([(2, 5, "=A1+B2+A1")])
    app.table.goto(2, 5)
    app.update()
    tags = [t for t in v.tag_names() if t.startswith("ref")]
    assert len(tags) == 2                                               # A1 twice = one colour
    assert v.tag_cget(tags[0], "foreground") == REF_COLORS[0]
    first = v.tag_ranges(tags[0])
    assert v.get(str(first[0]), str(first[1])) == "A1" and len(first) == 4          # both A1s in that colour
    second = v.tag_ranges(tags[1])
    assert v.tag_cget(tags[1], "foreground") == REF_COLORS[1] and v.get(str(second[0]), str(second[1])) == "B2"
