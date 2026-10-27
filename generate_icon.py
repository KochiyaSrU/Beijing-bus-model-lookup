"""Generate the Windows application icon for the bus model lookup tool."""

from PIL import Image, ImageDraw


SIZE = 512
BACKGROUND = "#103B66"
BUS_RED = "#C92936"
BUS_DARK = "#8E1D27"
WINDOW = "#DDEEFF"
ACCENT = "#F4C542"


def scale(value: int) -> int:
    return value * SIZE // 256


def rectangle(draw: ImageDraw.ImageDraw, bounds: tuple[int, int, int, int], **kwargs: object) -> None:
    draw.rounded_rectangle(tuple(scale(value) for value in bounds), radius=scale(18), **kwargs)


def main() -> None:
    image = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # Circular blue field makes the glyph recognizable on light and dark taskbars.
    draw.ellipse((scale(8), scale(8), scale(248), scale(248)), fill=BACKGROUND)
    rectangle(draw, (35, 70, 185, 177), fill=BUS_RED, outline=BUS_DARK, width=scale(5))
    draw.rounded_rectangle((scale(48), scale(87), scale(170), scale(129)), radius=scale(7), fill=WINDOW)
    for left in (55, 91, 127):
        draw.line((scale(left), scale(89), scale(left), scale(127)), fill=BUS_RED, width=scale(5))
    draw.rectangle((scale(51), scale(141), scale(168), scale(149)), fill=ACCENT)
    draw.ellipse((scale(57), scale(159), scale(83), scale(185)), fill="#1C2530")
    draw.ellipse((scale(138), scale(159), scale(164), scale(185)), fill="#1C2530")
    draw.ellipse((scale(63), scale(165), scale(77), scale(179)), fill="#E9EEF3")
    draw.ellipse((scale(144), scale(165), scale(158), scale(179)), fill="#E9EEF3")

    # A magnifier communicates that the application is a vehicle lookup tool.
    draw.ellipse((scale(146), scale(137), scale(210), scale(201)), fill="#FFFFFF", outline=ACCENT, width=scale(7))
    draw.line((scale(193), scale(187), scale(224), scale(218)), fill=ACCENT, width=scale(13))
    draw.line((scale(193), scale(187), scale(224), scale(218)), fill="#FFFFFF", width=scale(5))

    image.save(
        "北京公交车型查询器.ico",
        format="ICO",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )


if __name__ == "__main__":
    main()
