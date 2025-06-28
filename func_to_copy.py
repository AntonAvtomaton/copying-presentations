import win32com.client as win32
import find_dirs

def copy_slide_to_new_presentation(source_file_dir, output_file):
    """
    Функция для копирования слайдов
    :param source_file_dir: Путь к исходной презентации
    :param output_file: Новая презентация, куда будут сохраняться слайды
    """

    try:
        # Запускаем PowerPoint
        powerpoint = win32.gencache.EnsureDispatch("PowerPoint.Application")
        powerpoint.Visible = True

        # Открываем исходную презентацию
        source_pres = powerpoint.Presentations.Open(source_file_dir)

        #  Перебираем все слайды в исходной презентации. Count считает количество слайдов для наглядности
        count = 0
        for slide in source_pres.Slides:
            count += 1
            slide.Copy()  # Копируем в буфер обмена

            # Вставляем в новую презентацию (на первый слайд)
            output_file.Slides.Paste()
            print(f"Слайд {count} успешно скопирован в новую презентацию: {find_dirs.dir_output_file}")

        # Сохраняем новую презентацию
        output_file.Save()
        print("Презентация успешно сохранена")

    except Exception as e:
        print(f"Произошла ошибка: {e}")

    finally:
        # Закрываем презентации и выходим
        source_pres.Close()
        powerpoint.Quit()
