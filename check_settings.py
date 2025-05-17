"""Файл проверка объектов на соответствие требованиям"""
import base64
import json
import time
import os
from datetime import datetime

import urllib3
import requests
# import getpass

from reader_file_settings import all_settings_and_object_with_type
from writers import (
    write_all,
    write_policies_inconsistencies,
    write_tasks_inconsistencies
)


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


session = requests.Session()
NAME_SERVER = '192.168.56.101'
ADDRESS = f'https://{NAME_SERVER}:13299/api/v1.0'
URL_START = f'{ADDRESS}/Session.StartSession'
URL_DATA_POLICY = f'{ADDRESS}/Policy.GetPolicyData'
URL_DATA_TASKS = f'{ADDRESS}/Tasks.GetTaskData'
URL_CONTENT_POLICY = f'{ADDRESS}/Policy.GetPolicyContents'
URL_SECTIONS = f'{ADDRESS}/SsContents.SS_GetNames'
URL_READ_SECTION = f'{ADDRESS}/SsContents.Ss_Read'
URL_DATA_PROFILES = f'{ADDRESS}/PolicyProfiles.EnumProfiles'
URL_PROFILES_SETTINGS = f'{ADDRESS}/PolicyProfiles.GetProfileSettings'
ACTIVE = {True: 'Активировано', False: 'Деактивировано'}
MODELS = {'WSEE': 'KSWS', 'KES': 'KESW(12.0.0)', '1103': 'AGENT',
          'kesl': 'KESL', 'KESMAC12': 'KESMAC', 'KSVLA': 'KSVLA'}
TYPES_OF_TASKS = {
    'ScanserverUpdaterUpdaterTaskType': 'Обновление баз',
    'UpdateBases': 'Обновление баз программы',
    'UpdaterTaskType': 'Обновление',
    'updater': 'Обновление',
    'Update': 'Обновление',
    'Updater': 'Обновление',
    'OnDemand': 'Проверка по требованию',
    'ODS': 'Поиск вирусов',
    'group_ods': 'Поиск вирусов',
    'ods': 'Поиск вирусов',
    'KLNAG_TASK_VAPM_SEARCH': 'Поиск уязвимостей и требуемых обновлений'
}
INHERITED = {True: 'Унаследована', False: 'Не унаследована'}


def start_session() -> str:
    """Функция авторизации на сервере администрирования KSC."""
    # login = input('Введите логин\n')
    login = 'Администратор'
    # password = getpass.getpass('Введите пароль\n')
    password = 'Admin123'
    user64 = base64.b64encode(login.encode('utf-8')).decode("utf-8")
    password64 = base64.b64encode(password.encode('utf-8')).decode("utf-8")
    headers: dict = {
        'Authorization': f'KSCBasic user="{user64}", pass="{password64}"',
        'Content-Type': 'application/json',
    }
    data: dict = {}
    try:
        response = return_response(headers, URL_START, data)
        return json.loads(response.text)['PxgRetVal']
    except Exception as err:
        # Добавление собственного комментария к полученной ошибке с сервера
        raise ConnectionError(
            '\n\tНе удалось подключиться к серверу.\n'
            '\tОшибка в авторизации на сервере.\n'
        ) from err


def search_profiles(idx: int, policy: dict, model: str, headers: dict) -> list:
    """
    Функция, с помощью которой можно собирать
    информацию о профилях политик KES\n
    с сервера администрирования KSC.
    """
    parametrs = []
    data: dict = {"nPolicy": idx}
    response = return_response(headers, URL_DATA_PROFILES, data)
    description = json.loads(response.text)['PxgRetVal']
    if description.keys():
        profiles = list(description.keys())
        for profile in profiles:
            profile = '|'.join(
                [
                    str(policy['KLPOL_ID']),
                    str(policy['KLPOL_DN']),
                    str(policy['KLPOL_GROUP_NAME']),
                    str(profile),
                    str(model),
                    str(ACTIVE[description[profile]['value'][
                        'KLSSPOL_PRF_ENABLED']['value']['KLPRSS_Val']])
                ]
            )
            parametrs.append(profile)
    return parametrs


def search_policies(headers: dict, start_idx: int,
                    end_idx: int) -> tuple[list[str], list[str]]:
    """
    Функция, с помощью которой можно собирать информацию о политиках KES\n
    с сервера администрирования KSC.
    """
    list_policies = []
    list_profiles = []
    for idx in range(start_idx, end_idx):
        data: dict = {"nPolicy": idx}
        response = return_response(headers, URL_DATA_POLICY, data)
        if 'PxgRetVal' in response.text:
            description = json.loads(response.text)['PxgRetVal']
            model = MODELS.get(
                description['KLPOL_PRODUCT'],
                'Не удалось определить тип программы'
            )
            parameters = '|'.join(
                [
                    str(description['KLPOL_ID']),
                    str(description['KLPOL_DN']),
                    str(description['KLPOL_GROUP_NAME']),
                    str(model),
                    str(INHERITED[description['KLPOL_INHERITED']]),
                    str(ACTIVE[description['KLPOL_ACTIVE']])
                ]
            )
            list_policies.append(parameters)
            profiles = search_profiles(idx, description, model, headers)
            if profiles:
                list_profiles.extend(profiles)
    return list_policies, list_profiles


def search_tasks(headers: dict, start_idx: int, end_idx: int) -> list:
    """
    Функция, с помощью которой можно собирать информацию о задачах KES\n
    с сервера администрирования KSC.
    """
    list_tasks: list = []
    for idx in range(start_idx, end_idx):
        data = {"strTask": f'{idx}'}
        response = return_response(headers, URL_DATA_TASKS, data)
        if 'PxgRetVal' in response.text:
            description = json.loads(response.text)['PxgRetVal']
            group = description['TASK_INFO_PARAMS']['value'].get(
                'PRTS_TASK_GROUPNAME',
                'Задачи для наборов устройств'
            )
            model = MODELS.get(
                description['TASKID_PRODUCT_NAME'],
                'Не удалось определить тип программы'
            )
            type_task = TYPES_OF_TASKS.get(
                description['TASK_NAME'],
                'Данный тип не поддерживается скриптом'
            )
            parameters = '|'.join(
                [
                    description['TASK_UNIQUE_ID'],
                    description['TASK_INFO_PARAMS']['value']['DisplayName'],
                    group,
                    model,
                    type_task
                ]
            )
            list_tasks.append(parameters)
    return list_tasks


def download_chapter_from_policies(headers: dict, object_idx: int,
                                   section: str):
    """
    Данная функция-инструмент, позволяет выгружать\n
    раздел параметров с политики KES через API c KSC.
    - Инструкция по данной функции приведена в файле instrumentarium.py
    """
    data: dict = {'nPolicy': object_idx}
    response = return_response(headers, URL_DATA_POLICY, data)
    description = json.loads(response.text)['PxgRetVal']
    data = {
        'nPolicy': description['KLPOL_ID'],
        'nRevisionId': 0,
        'nLifeTime': 0
    }
    response = return_response(headers, URL_CONTENT_POLICY, data)
    content = json.loads(response.text)['PxgRetVal']
    data = {
        'wstrID': content,
        'wstrProduct': description['KLPOL_PRODUCT'],
        'wstrVersion': description['KLPOL_VERSION'],
        'wstrSection': section
    }
    response = return_response(headers, URL_READ_SECTION, data)
    settings = json.loads(response.text)
    return settings


def download_chapter_from_profiles(headers: dict, object_idx: int,
                                   profile: str, section: str):
    """
    Данная функция-инструмент, позволяет выгружать\n
    раздел параметров с профиля KES через API c KSC.
    """
    data: dict = {'nPolicy': object_idx}
    response = return_response(headers, URL_DATA_POLICY, data)
    description = json.loads(response.text)['PxgRetVal']
    data = {
        'nPolicy': description['KLPOL_ID'],
        'szwName': profile,
        'nRevisionId': 0,
        'nLifeTime': 0
    }
    response = return_response(headers, URL_PROFILES_SETTINGS, data)
    content = json.loads(response.text)['PxgRetVal']
    data = {
        'wstrID': content,
        'wstrProduct': description['KLPOL_PRODUCT'],
        'wstrVersion': description['KLPOL_VERSION'],
        'wstrSection': section
    }
    response = return_response(headers, URL_READ_SECTION, data)
    settings = json.loads(response.text)
    return settings


def download_setting_in_task(headers: dict, object_idx: int):
    """
    Данная функция-инструмент, позволяет находить\n
    настройки параметров в задаче KES.
    """
    data = {'strTask': f'{object_idx}'}
    response = return_response(headers, URL_DATA_TASKS, data)
    description = json.loads(response.text)
    return description


def reload_data_policies(headers: dict, objects: dict):
    """
    Функция получения информации с API для политики
    """
    # для политики, собираем информацию с KSC с раздела,
    # для параметров которого включен режит Аудит
    for idx, object_idx in enumerate(objects['Индекс']):
        # определяем базовый тип проверяемой политики
        check_type: str = objects['Тип проверки'][idx]
        if '_' not in check_type:
            base_check_type = check_type
        else:
            base_check_type = check_type[:check_type.find('_')]
        searchs_sections: dict[str, dict] = {base_check_type: {}}
        # по типу проверяемой политики собираем данные
        # с необходимых разделов
        for section in TYPES[base_check_type]['sections']:
            # получаем настройки с раздела
            settings = download_chapter_from_policies(
                headers, object_idx, section)
            searchs_sections[base_check_type][section] = searchs_sections[
                base_check_type].get(section, settings)
        # к политики мы прибавляем data-JSON настроек в текущей политике
        objects['Настройки'][idx] = searchs_sections[base_check_type]
    return objects


def reload_data_profiles(headers: dict, objects: dict):
    """
    Функция получения информации с API для профиля
    """
    for idx, object_idx in enumerate(objects['Индекс']):
        check_type: str = objects['Тип проверки'][idx]
        profile: str = objects['Профиль'][idx]
        if '_' not in check_type:
            base_check_type = check_type
        else:
            base_check_type = check_type[:check_type.find('_')]
        searchs_sections: dict[str, dict] = {base_check_type: {}}
        for section in TYPES[base_check_type]['sections']:
            settings = download_chapter_from_profiles(
                headers, object_idx, profile, section)
            searchs_sections[base_check_type][section] = searchs_sections[
                base_check_type].get(section, settings)
        objects['Настройки'][idx] = searchs_sections[base_check_type]
    return objects


def reload_data_tasks(headers: dict, objects: dict):
    """
    Функция получения информации с API для задачи
    """
    for idx, object_idx in enumerate(objects['Индекс']):
        settings = download_setting_in_task(headers, object_idx)
        objects['Настройки'][idx] = settings
    return objects


def _extract_value(current_value):
    """Возвращает значение, приведённое к нужному виду для сравнения."""
    if isinstance(current_value, list):
        return current_value[0] if current_value else tuple(current_value)
    if isinstance(current_value, int) and current_value < 0:
        current_value = str(current_value)
        return current_value
    if isinstance(current_value, dict):
        return current_value.get('value', current_value)
    return current_value


def _search_value_checkbox(**kwargs):
    """
    Универсальная функция проверки флажка настройки.\n
    Возвращает:
      - 'Нет значения в JSON' если отсутствуют данные,
      - 'Соответствует' если значение найдено в словаре или через функцию,
      - или сообщение об ошибке при несоответствии.
    """
    result = 'Нет значения в JSON'
    try:
        nested_values = settings_base_type_verif[
            kwargs['check_type']][
                kwargs['chapter']][
                    kwargs['parametr_name']][
                        kwargs['setting_name']]
        # если в вложенных значениях допущена ошибка, то возврат
        if not nested_values:
            return result
        section = nested_values.get('Раздел в JSON c KSC', False)
        json_path = nested_values['Маршрут в JSON c KSC']
        blocked_path = nested_values.get('Маршрут в JSON до блокировки', False)
        if kwargs['value_field'] == 'KLPRSS_Mnd' and blocked_path:
            json_path = blocked_path
        # раздел в API KSC
        settings_section = kwargs['objects']['Настройки'][kwargs['idx']]
        # текущее значение
        if kwargs['value_field'] and section:
            current_value = get_value(
                settings_section[section], json_path)[kwargs['value_field']]
        else:
            current_value = get_value(settings_section, json_path)
    except (KeyError, IndexError, TypeError):
        # если путь json_path указан неверно, или в процессе появилась ошибка
        # то возврат
        return result
    # проверка того, что требумое значение является словарем
    # или запрет на редактирование
    result_key = kwargs['result_key']
    required_value = kwargs['required_value']
    nested = nested_values[result_key]
    value = _extract_value(current_value)
    if isinstance(nested, dict):
        check_value = nested[required_value]
        # Для особых ключей с хешами — используем .get
        result = check_value.get(value, 'Соответствует')

    elif callable(nested):
        try:
            check_value = nested({required_value: 'Соответствует'})
            result = check_value[value]
        except (KeyError, TypeError):
            result = f'Не соответствует!\nУстановлено: {value}'
    return result


def _value_verification_area(**kwargs):
    """
    Универсальная функция проверки для "Области защиты".\n
    Возвращает:
      - 'Нет значения в JSON' если отсутствуют данные,
      - 'Соответствует' если значение найдено,
      - или сообщение об ошибке при несоответствии.
    """
    value_result = 'Нет значения в JSON'
    try:
        nested_values = settings_base_type_verif[
            kwargs['check_type']][
                kwargs['chapter']][
                    kwargs['parametr_name']][
                        kwargs['setting_name']]
        # если в вложенных значениях допущена ошибка, то возврат
        if not nested_values:
            return value_result
        section = nested_values.get('Раздел в JSON c KSC', False)
        path_all_value = nested_values['Маршрут в JSON c KSC']
        keys = nested_values['Ключи']
        settings_section = kwargs['objects']['Настройки'][kwargs['idx']]
        current_value = None
        bad_values = []
        if kwargs['value_field'] and section:
            all_value = get_value(
                settings_section[section], path_all_value)[
                    kwargs['value_field']]
        else:
            all_value = get_value(settings_section, path_all_value)
        for _, value in enumerate(all_value):
            if get_value(value, keys['path_type']) == keys['value_type']:
                current_value = get_value(value, keys['wanted_value_path'])
                if current_value not in kwargs['required_value'].split(', '):
                    activate = keys.get('activate', False)
                    if activate:
                        activate = get_value(value, keys['activate'])
                    if activate:
                        bad_values.append(f'{current_value}({keys[activate]})')
                elif current_value in kwargs['required_value'].split(', '):
                    activate = keys.get('activate', False)
                    if activate:
                        activate = get_value(value, keys['activate'])
                    if not activate:
                        bad_values.append(f'{current_value}({keys[activate]})')

        lookup = nested_values.get(
            'Требуемое значение настройки', False)
        if lookup:
            value_result = lookup[kwargs['required_value']].get(
                current_value, False)
        if kwargs['required_value'] not in ['ON', 'OFF']:
            if bad_values:
                value_result = (
                    'Не соответствует!\n'
                    f"Обнаружено: \n"
                    f"{', '.join(bad_values)}"
                )
            else:
                value_result = 'Соответствует'
    except (KeyError, IndexError, TypeError):
        # если путь json_path указан неверно, или в процессе появилась ошибка
        # то возврат
        return value_result
    return value_result


def _configuration_fields(config: dict, base_type_verif: str,
                          type_verif: str, setting_idx: int, fl_ts=False):
    """
    Формирует вспомогательные ключи (поля), для упрощения читаемости кода
    """
    if fl_ts:
        return {
            'check_type': base_type_verif,
            'chec_type_exp': type_verif,
            'chapter': config['Раздел'][setting_idx],
            'module': config['Модуль'][setting_idx],
            'parameter_name': config['Название параметра'][setting_idx],
            'setting_name': config['Настройка KSC'][setting_idx],
            'standard_section': config['Пункт стандарта'][setting_idx],
            'standard_requirement': config[
                'Требование стандарта'][setting_idx],
            'required_value': config[
                'Требуемое значение настройки'][setting_idx],
            'exception_flag': config['Исключение'][setting_idx],
            'except_value': config.get(type_verif, '')
        }
    return {
        'check_type': base_type_verif,
        'chec_type_exp': type_verif,
        'chapter': config['Раздел'][setting_idx],
        'module': config['Модуль'][setting_idx],
        'parameter_name': config['Название параметра'][setting_idx],
        'setting_name': config['Настройка KSC'][setting_idx],
        'standard_section': config['Пункт стандарта'][setting_idx],
        'standard_requirement': config['Требование стандарта'][setting_idx],
        'required_value': config['Требуемое значение настройки'][setting_idx],
        'required_block': config['Запрет на редактирование'][setting_idx],
        'exception_flag': config['Исключение'][setting_idx],
        'except_value': config.get(type_verif, ''),
        'except_block': config.get(f'BLOCK {type_verif}', '')
    }


def return_results_check(objects: dict, object_idx: int,
                         setting: dict, setting_idx: int):
    """
    Выполняет проверку настройки и запрета редактирования.
    """
    if setting['parameter_name'] != 'Область защиты':
        support_function = _search_value_checkbox
    else:
        support_function = _value_verification_area
    # Базовые результаты
    # результат проверки на запрет редактирования у необходимой настройки
    block_result = _search_value_checkbox(
        check_type=setting['check_type'],
        chapter=setting['chapter'],
        parametr_name=setting['parameter_name'],
        setting_name=setting['setting_name'],
        objects=objects,
        idx=object_idx,
        required_value=setting['required_block'],
        result_key='Запрет на редактирование',
        value_field='KLPRSS_Mnd'
    )
    # результат проверки значения необходимой настройки
    value_result = support_function(
        check_type=setting['check_type'],
        chapter=setting['chapter'],
        parametr_name=setting['parameter_name'],
        setting_name=setting['setting_name'],
        objects=objects,
        idx=object_idx,
        required_value=setting['required_value'],
        result_key='Требуемое значение настройки',
        value_field='KLPRSS_Val'
    )
    # если для настрокий допущено исключение, то
    if value_result != 'Соответствует':
        # Исключение для значения
        if setting['exception_flag'] == 'ON':
            try:
                setting['required_value'] = setting[
                    'except_value'][setting_idx]
                value_result = support_function(
                    check_type=setting['check_type'],
                    chapter=setting['chapter'],
                    parametr_name=setting['parameter_name'],
                    setting_name=setting['setting_name'],
                    objects=objects,
                    idx=object_idx,
                    required_value=setting['required_value'],
                    result_key='Требуемое значение настройки',
                    value_field='KLPRSS_Val'
                )
            except IndexError:
                value_result = (
                    f'{value_result}\n'
                    f'Для типа «{setting["chec_type_exp"]}» '
                    'не определенно исключений!'
                )
    # Исключение для флага блокировки
    if block_result != 'Соответствует':
        if setting['exception_flag'] == 'ON':
            try:
                setting['required_block'] = setting[
                    'except_block'][setting_idx]
                block_result = _search_value_checkbox(
                    check_type=setting['check_type'],
                    chapter=setting['chapter'],
                    parametr_name=setting['parameter_name'],
                    setting_name=setting['setting_name'],
                    objects=objects,
                    idx=object_idx,
                    required_value=setting['required_block'],
                    result_key='Запрет на редактирование',
                    value_field='KLPRSS_Mnd'
                )
            except IndexError:
                block_result = (
                    f'{block_result}\n'
                    f'Для типа «{setting["chec_type_exp"]}» '
                    'не определенно исключений!'
                )
    return {'value_result': value_result, 'block_result': block_result}


def return_results_check_tasks(objects: dict, object_idx: int,
                               setting: dict, setting_idx: int):
    """
    Выполняет проверку настройки и запрета редактирования.
    """
    if (
        setting['parameter_name'] != 'Область проверки' and
        'обновлений' not in setting['parameter_name']
    ):
        support_function = _search_value_checkbox
    else:
        support_function = _value_verification_area
    # результат проверки значения необходимой настройки
    value_result = support_function(
        check_type=setting['check_type'],
        chapter=setting['chapter'],
        parametr_name=setting['parameter_name'],
        setting_name=setting['setting_name'],
        objects=objects,
        idx=object_idx,
        required_value=setting['required_value'],
        result_key='Требуемое значение настройки',
        value_field=False
    )
    # если для настрокий допущено исключение, то
    if value_result != 'Соответствует':
        # Исключение для значения
        if setting['exception_flag'] == 'ON':
            try:
                setting['required_value'] = setting[
                    'except_value'][setting_idx]
                value_result = support_function(
                    check_type=setting['check_type'],
                    chapter=setting['chapter'],
                    parametr_name=setting['parameter_name'],
                    setting_name=setting['setting_name'],
                    objects=objects,
                    idx=object_idx,
                    required_value=setting['required_value'],
                    result_key='Требуемое значение настройки',
                    value_field=False
                )
            except IndexError:
                value_result = (
                    f'{value_result}\n'
                    f'Для типа «{setting["chec_type_exp"]}» '
                    'не определенно исключений!'
                )
    return {'value_result': value_result}


def formation_mismatch_string(objects: dict, object_idx: int,
                              setting: dict, result: dict[str, str],
                              fl_ts=False, fl_pf=False):
    """Добавляет одну запись о несоответствии в результат."""
    navigation = '.'.join([
        setting['chapter'], setting['module'], setting['parameter_name']
    ]).replace('.', ' → ').replace(' → Модуль', '').replace(
        ' → Основные параметры', '')
    if fl_ts:
        return '|'.join(
            [
                str(objects['Название'][object_idx]),
                str(objects['Группа'][object_idx]),
                str(setting['standard_section']),
                str(setting['standard_requirement']),
                str(navigation),
                str(setting['setting_name']),
                str(setting['required_value']),
                str(result['value_result']),
            ]
        )
    if fl_pf:
        return '|'.join(
            [
                str(objects['Название'][object_idx]),
                str(objects['Группа'][object_idx]),
                str(objects['Профиль'][object_idx]),
                str(setting['standard_section']),
                str(setting['standard_requirement']),
                str(navigation),
                str(setting['setting_name']),
                str(setting['required_block']),
                str(setting['required_value']),
                str(result['block_result']),
                str(result['value_result']),
            ]
        )
    return '|'.join(
        [
            str(objects['Название'][object_idx]),
            str(objects['Группа'][object_idx]),
            str(setting['standard_section']),
            str(setting['standard_requirement']),
            str(navigation),
            str(setting['setting_name']),
            str(setting['required_block']),
            str(setting['required_value']),
            str(result['block_result']),
            str(result['value_result']),
        ]
    )


def check_up_policies(settings: dict, objects: dict, foldername: str):
    """Проверяет параметры политики на соответствие стандартам."""
    all_objects_incongruity = []
    for object_idx in range(len(objects['Индекс'])):
        objects_incongruity = []
        # type_verif - это тип которые изначально указан у политики в
        # в поле 'Тип проверки'
        type_verif = objects['Тип проверки'][object_idx]
        check_sheet_name = objects['Чек-лист'][object_idx]
        # настройки из листа с проверками
        config = settings[check_sheet_name]
        if '_' not in type_verif:
            # Это лист, в котоором содержатся все проверяемые пункты проверки
            base_type_verif = type_verif
        else:
            base_type_verif = type_verif[:type_verif.find('_')]
        # проходимся по имеющимся включенным проверкам, по разделам
        for setting_idx in range(len(config['Раздел'])):
            # строка проверки, по которой проверяем пролитику на соответсвие
            setting = _configuration_fields(config, base_type_verif,
                                            type_verif, setting_idx)
            # результат (требуемое значение и запрет редактирования)
            result = return_results_check(objects, object_idx,
                                          setting, setting_idx)
            if (
                result['value_result'] != 'Соответствует' or
                result['block_result'] != 'Соответствует'
            ):
                bad_incongruity = formation_mismatch_string(
                    objects, object_idx, setting, result)
                all_objects_incongruity.append(bad_incongruity)
            incongruity = formation_mismatch_string(
                objects, object_idx, setting, result)
            objects_incongruity.append(incongruity)
        os.makedirs(fr'{foldername}\policies', exist_ok=True)
        write_policies_inconsistencies(
            result=objects_incongruity,
            filename=fr"{foldername}\policies\{
                objects['Группа'][object_idx]}-{
                    objects['Название'][object_idx]}.xlsx"
        )
    return all_objects_incongruity


def check_up_profiles(settings: dict, objects: dict, foldername: str):
    """Проверяет параметры политики на соответствие стандартам."""
    all_objects_incongruity = []
    for object_idx in range(len(objects['Индекс'])):
        objects_incongruity = []
        # type_verif - это тип которые изначально указан у политики в
        # в поле 'Тип проверки'
        type_verif = objects['Тип проверки'][object_idx]
        check_sheet_name = objects['Чек-лист'][object_idx]
        # настройки из листа с проверками
        config = settings[check_sheet_name]
        if '_' not in type_verif:
            # Это лист, в котоором содержатся все проверяемые пункты проверки
            base_type_verif = type_verif
        else:
            base_type_verif = type_verif[:type_verif.find('_')]
        # проходимся по имеющимся включенным проверкам, по разделам
        for setting_idx in range(len(config['Раздел'])):
            # строка проверки, по которой проверяем пролитику на соответсвие
            setting = _configuration_fields(config, base_type_verif,
                                            type_verif, setting_idx)
            # результат (требуемое значение и запрет редактирования)
            result = return_results_check(objects, object_idx,
                                          setting, setting_idx)
            if (
                result['value_result'] != 'Соответствует' or
                result['block_result'] != 'Соответствует'
            ):
                bad_incongruity = formation_mismatch_string(
                    objects, object_idx, setting, result, fl_pf=True)
                all_objects_incongruity.append(bad_incongruity)
            incongruity = formation_mismatch_string(
                objects, object_idx, setting, result, fl_pf=True)
            objects_incongruity.append(incongruity)
        os.makedirs(fr'{foldername}\profiles', exist_ok=True)
        write_policies_inconsistencies(
            result=objects_incongruity,
            filename=fr"{foldername}\profiles\{
                objects['Группа'][object_idx]}-{
                    objects['Название'][object_idx]}-{
                        objects['Профиль'][object_idx]
                    }.xlsx",
            fl_pf=True
        )
    return all_objects_incongruity


def check_up_tasks(settings: dict, objects: dict, foldername: str):
    """Проверяет параметры задачи на соответствие стандартам."""
    all_objects_incongruity = []
    for object_idx in range(len(objects['Индекс'])):
        objects_incongruity = []
        type_verif = objects['Тип проверки'][object_idx]
        check_sheet_name = objects['Чек-лист'][object_idx]
        config = settings[check_sheet_name]
        if '_' not in type_verif or type_verif.count('_') == 1:
            base_type_verif = type_verif
        else:
            base_type_verif = type_verif[:type_verif.rfind('_')]
        for setting_idx in range(len(config['Раздел'])):
            # строка проверки, по которой проверяем пролитику на соответсвие
            setting = _configuration_fields(config, base_type_verif,
                                            type_verif, setting_idx,
                                            fl_ts=True)
            result = return_results_check_tasks(objects, object_idx,
                                                setting, setting_idx)
            if result['value_result'] != 'Соответствует':
                bad_incongruity = formation_mismatch_string(
                    objects, object_idx, setting, result, fl_ts=True)
                all_objects_incongruity.append(bad_incongruity)
            incongruity = formation_mismatch_string(
                objects, object_idx, setting, result, fl_ts=True)
            objects_incongruity.append(incongruity)
        os.makedirs(fr'{foldername}\tasks', exist_ok=True)
        write_tasks_inconsistencies(
            result=objects_incongruity,
            filename=fr"{foldername}\tasks\{
                objects['Группа'][object_idx]}-{
                    objects['Название'][object_idx]}.xlsx"
        )
    return all_objects_incongruity


def main():
    """Функция управления остальными функциями."""
    token = start_session()
    date_created = datetime.now().strftime("%Y-%m-%d")
    base_path = os.path.join("отчеты")
    folder = create_unique_date_folder(date_created, base_path)
    headers = {'X-KSC-Session': token, 'Content-Type': 'application/json'}
    start = time.perf_counter()
    inconsistencies = {
        'policies_inconsistencies': {
            'data': [],
            'sheetname': 'Наблюдения по политикам',
            'counter': {'new': 0, 'delete': 0, 'all': 0},
            'start_sheet': 'policies',
            'func_check_up': check_up_policies,
            'func_reload': reload_data_policies
        },
        'profiles_inconsistencies': {
            'data': [],
            'sheetname': 'Наблюдения по профилям',
            'counter': {'new': 0, 'delete': 0, 'all': 0},
            'start_sheet': 'profiles',
            'func_check_up': check_up_profiles,
            'func_reload': reload_data_profiles
        },
        'tasks_inconsistencies': {
            'data': [],
            'sheetname': 'Наблюдения по задачам',
            'counter': {'new': 0, 'delete': 0, 'all': 0},
            'start_sheet': 'tasks',
            'func_check_up': check_up_tasks,
            'func_reload': reload_data_tasks
        }
    }
    flag = '_' not in folder
    # Записывааются данные в файлы
    for name, data in inconsistencies.items():
        settings, objects = all_settings_and_object_with_type(
            sheetname=data['start_sheet'], flag=False)
        if objects:
            objects = data['func_reload'](headers, objects)
            data['data'] = data['func_check_up'](
                    settings, objects, folder)
            data['data'], data['counter'] = comparison_inconsistencies(
                data['data'], flag, name
            )
        print(name, data['counter'])
    objects = {
        'policies': {
            'data': [],
            'sheetname': 'Политики',
            'counter': {'new': 0, 'delete': 0, 'all': 0}
        },
        'profiles': {
            'data': [],
            'sheetname': 'Профили',
            'counter': {'new': 0, 'delete': 0, 'all': 0},
        },
        'tasks': {
            'data': [],
            'sheetname': 'Задачи',
            'counter': {'new': 0, 'delete': 0, 'all': 0}
        }
    }
    activates = {
        'policies_activate': {
            'data': [],
            'sheetname': 'Изменение в активности политик',
            'status': ''
        },
        'profiles_activate': {
            'data': [],
            'sheetname': 'Изменение в активности профилей',
            'status': ''
        }
    }
    objects['policies']['data'], objects['profiles']['data'] = search_policies(
        headers=headers, start_idx=1, end_idx=200)
    objects['policies']['data'], objects['policies']['counter'], activates[
        'policies_activate']['data'], activates[
            'policies_activate']['status'] = news_objects(
        objects['policies']['data'], flag, 'policies'
    )
    print('policies', objects['policies']['counter'])
    objects['profiles']['data'], objects['profiles']['counter'], activates[
        'profiles_activate']['data'], activates[
            'profiles_activate']['status'] = news_objects(
        objects['profiles']['data'], flag, 'profiles'
    )
    print('profiles', objects['profiles']['counter'])
    objects['tasks']['data'] = search_tasks(headers=headers, start_idx=1,
                                            end_idx=200)
    objects['tasks']['data'], objects['tasks']['counter'], _, _ = news_objects(
        objects['tasks']['data'], flag, 'tasks'
    )
    print('tasks', objects['tasks']['counter'])
    result = inconsistencies | objects | activates
    dump_json('result.json', result)
    write_all(
        result=result,
        filename=fr'{folder}\несоответсвия.xlsx'
    )
    end = time.perf_counter()
    print(f"Время выполнения: {end - start:.6f} секунд")


if __name__ == '__main__':
    main()
