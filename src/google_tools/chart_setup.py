from PIL import ImageColor

def hex_to_rgb(hex_code):
    r, g, b = [c / 256 for c in ImageColor.getcolor(hex_code, 'RGB')]
    out = {'red':r, 'green':g, 'blue':b}
    return out
