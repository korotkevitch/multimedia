from PIL import Image
import os
from moviepy.editor import *
import glob

path_name_images = r'D:\Cardsmaker\clients\one_gif_holidays\target'

path_source = r'D:\Cardsmaker\clients\one_gif_holidays\target\yana'
path_target = r'D:\Cardsmaker\clients\one_gif_video\target\Яна'

# Программа проходит по списку всех картинок в папке и делает для них resize и crop.
# При этом для файлов gif, чтобы не пропала анимация, используется VideoFileClip, а для остальных ImageClip.
# gif-файлы сохраняем как gif, а остальные - png

# создаем список файлов для последующей обработки
q = os.listdir(path_source)

# считаем кол-во файлов без учета папок
count_files = len([f for f in q if os.path.isfile(os.path.join(path_source, f))])

for i in range(count_files):
    file = q[i]
    file_body, file_ext = os.path.splitext(file)

    if file_ext == '.gif':
        img = VideoFileClip(os.path.join(path_source, file))

        old_width = img.size[0]
        old_height = img.size[1]

        # задаем требуемый размер: высоту для гориз. фото или ширину для вертикальных

        max_size = 500

        if old_width > old_height:  # для горизонтальных фото

            wpercent = float(old_height / max_size)
            new_width = int(old_width / wpercent)
            img = img.resize((int(new_width), int(max_size)))

        else:  # для вертикального фото

            wpercent = float(old_width / max_size)
            new_height = int(old_height / wpercent)
            img = img.resize((int(max_size), int(new_height)))

        width, height = img.size  # Get dimensions

        left = (width - max_size) / 2
        top = (height - max_size) / 2
        right = (width + max_size) / 2
        bottom = (height + max_size) / 2

        # вырезаем 500*500 по центру
        img = img.crop(left, top, right, bottom)
        img.write_gif(os.path.join(path_target, file))


    else:
        img = Image.open(os.path.join(path_source, file))

        old_width = img.size[0]
        old_height = img.size[1]

# задаем требуемый размер: высоту для гориз. фото или ширину для вертикальных

        max_size = 500

        if old_width > old_height:  # для горизонтальных фото

            wpercent = float(old_height/max_size)
            new_width = int(old_width/wpercent)
            img = img.resize((int(new_width), int(max_size)))

        else:  # для вертикального фото

            wpercent = float(old_width/max_size)
            new_height = int(old_height/wpercent)
            img = img.resize((int(max_size), int(new_height)))


        width, height = img.size   # Get dimensions

        left = (width - max_size)/2
        top = (height - max_size)/2
        right = (width + max_size)/2
        bottom = (height + max_size)/2

# вырезаем 500*500 по центру
        img = img.crop((left, top, right, bottom))
        img.save(os.path.join(path_target, file_body + file_ext))