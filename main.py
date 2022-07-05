import xlwings
from PIL import Image
from time import sleep
import pyautogui

fps = 30
wide = 80
black = (0, 0, 0)
white = (255, 255, 255)

wb = xlwings.Book('Bad_apple.xlsx')
sheet = wb.sheets['Planilha1']


def resize_gray(image, wdt=72):
    width, height = image.size
    aspect = height/width
    hgt = int(aspect*wdt)
    res_gray = image.resize((wdt, hgt)).convert('L')
    bw = res_gray.point(lambda x: 0 if x < 128 else 255, '1')
    return bw


def draw_frame(image, wdt=wide+1):
    pixels = image.getdata()
    line = 1
    colun = 1
    for pixel in pixels:
        if colun == wdt:
            line += 1
            colun = 1

        if pixel > 128:
            sheet.cells(line, colun).color = white
        elif pixel < 128:
            sheet.cells(line, colun).color = black

        colun += 1


sleep(2)

try:

    for index in range(1, 6573):
        frame = Image.open(f'BadApple-frames/frame{index}.jpg')
        draw_frame(resize_gray(frame, wide))
        screen_shot = pyautogui.screenshot()
        screen_shot.save(f'Excel-screenshots/shot{index}.jpg')
        wb.sheets[1].name = f'{int(index/(6572*(1/100)))}%'


except KeyboardInterrupt:
    print('process finished')

wb.sheets[1].name = 'concluído!'
print('finished!')
