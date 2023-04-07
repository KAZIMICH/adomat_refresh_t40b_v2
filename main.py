import win32com.client
import os
import time
import sys
import data

file_list_vk = []
path_list_vk = []
exception_files = []
path_list_nx = [data.file_path_importNX, data.file_path_NXPOWER, data.file_path_NXSIGNAL]
path_db_2x = [data.file_path_db, data.file_path_db]


''' Для корректной работы приложения необходимо, чтобы во всех файлах экселя
    для всех запросов была снята галочка "фоновое обновление" '''


def name_and_path_vk(i):
    for root, dirs, files in os.walk(i):
        for file in files:
            if file.endswith('ВК.xlsx') and '~$' not in file \
                    or file.endswith('ВК.xls') and '~$' not in file \
                    or file.endswith('ВК.xlsm') and '~$' not in file:
                file_list_vk.append(os.path.join(file))
                path_list_vk.append(os.path.join(root, file))

    if len(file_list_vk) > 0:
        print(f'В папке с проектом найдено {len(file_list_vk)} файлов ВК для обновления')
        answer = dialog_yes_no('Вывести список файлов ВК для обновления?\nВведите Y или N')
        if answer == 'y':
            print_list(file_list_vk)
            print('_' * 100)
        else:
            print('_' * 100)
    else:
        print(input('Список пуст. Что-то пошло не так! Перезапустите программу...'))
        sys.exit()
    return file_list_vk, path_list_vk


def refresh_files(i, list_name, visible=False):
    if list_name != 'Файл БД':
        check_double(i, list_name)
    else:
        pass
    print(f'Файлы =={list_name}== готовы к обновлению')
    user_answer = dialog_yes_no('Обновить файлы?\nВведите Y или N')

    if user_answer == 'y':
        excel = win32com.client.DispatchEx("Excel.Application")
        counter = 1
        for j in i:
            file_check(j)
            print(f'Обновление файла {counter} =={list_name}==')
            print(j)
            wb = excel.Workbooks.Open(j)
            wb.Application.DisplayAlerts = visible
            wb.Application.EnableEvents = visible
            wb.Application.ScreenUpdating = visible
            wb.Application.Interactive = visible
            excel.Visible = visible
            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()
            time.sleep(1)
            wb.Save()
            wb.Close()
            excel.Quit()
            print('Файл', counter, 'отработан')
            counter += 1
        print(f'Отработано {len(i)} файлов в списке =={list_name}==')
        print(f'{len(exception_files)} файлов из списка =={list_name}== исключены из обновления:')
        print_list(exception_files)
        exception_files.clear()
        print('_' * 100)
    else:
        print('_' * 100)
        print('Выполнение приложения прервано пользователем')
        endTime = time.time()
        totalTime = endTime - startTime
        print('Работа завершена')
        print(input(f'Затраченное время = {int(totalTime)} секунд\n'))
        sys.exit()


def check_double(i, list_name):
    print(f'Проверка файлов =={list_name}== на дубли...')
    visited = set()
    dup = [x for x in i if x in visited or (visited.add(x) or False)]
    if len(dup) > 0:
        for j in dup:
            print(f'Документ {j} присутствует несколько раз')
        print(input('Удалите неактуальный файл и перезапустите приложение'))
        sys.exit()
    else:
        print('Проверка завершена. Дублей файлов не найдено.')
        print('_' * 100)


def file_check(file):
    valid = False
    while not valid:
        if os.path.exists(file):
            try:
                os.rename(file, file)
                valid = True
            except IOError:
                valid_flag = False
                print(f'Файл {file} открыт')
                answer = dialog_yes_no('Исключить этот файл из обновления?\nВведите Y или N')
                if answer == 'y':
                    exception_files.append(file)
                    return exception_files
                else:
                    print(input(f'Файл {file} открыт.\nДля принятия текущих изменений в файле сохраните его и нажмите '
                                f'Enter.\n'))
                valid = False
        else:
            print('Файл', file, 'был перемещен или удален')
            print(input('Для подтверждения нажмите Enter.\n'))
            valid = False
    return valid


def dialog_yes_no(i):
    while True:
        answers = {'yes': 1, 'y': 1, 'no': 0, 'n': 0}
        print(i)
        answer = input().lower()
        if answer in answers:
            return answer
        else:
            print('Ожидалось Y или N')


def print_list(i):
    for j in i:
        print(j)


if __name__ == '__main__':
    startTime = time.time()
    name_and_path_vk(data.path_folder)
    refresh_files(path_db_2x, 'Файл БД', True)

    refresh_files(path_list_vk, 'Файлы ВК')

    refresh_files(path_list_nx, 'Файлы NX', False)

    refresh_files(path_db_2x, 'Файл БД', True)

    endTime = time.time()
    totalTime = endTime - startTime
    print('Программа завершена')
    print(input(f'Затраченное время = {int(totalTime)} секунд\nНажмите Enter'))
    sys.exit()

