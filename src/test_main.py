from src import main
import pytest
import openpyxl


def test_add_file_postfix():
    name = main.add_file_postfix('no_postfix')
    assert name == 'no_postfix.xlsx'

    name = main.add_file_postfix('with_postfix.xlsx')
    assert name == 'with_postfix.xlsx'


def test_old_excel_exception():
    with pytest.raises(Exception):
        main.add_file_postfix('old.xls')


def test_parse_food_items():
    excel = openpyxl.load_workbook('商家資料填寫表格.xlsx')
    item_sheet = excel['品項']
    food_items, revere_index = main.parse_food_items(item_sheet)

    # assert exist
    assert food_items is not None
    assert revere_index is not None

    # assert kind
    assert food_items['0']['kindname'] == '風味炒飯'
    assert food_items['1']['kindname'] == '風味炒麵'
    assert food_items['2']['kindname'] == '燴飯(附蛋)'
    assert food_items['3']['kindname'] == '麻婆豆腐系列'
    assert food_items['4']['kindname'] == '日式咖哩飯'
    with pytest.raises(KeyError):
        food_items['5']

    # assert kind items
    assert food_items['0']['items']['0_0']['foodname'] == '風味蛋炒飯'
    assert food_items['0']['items']['0_0']['price'] == 50
    assert food_items['0']['items']['0_0']['memo'] is None
    assert food_items['1']['items']['1_6']['foodname'] == 'XO醬雙蛋(推薦)'
    assert food_items['1']['items']['1_6']['price'] == 70
    assert food_items['1']['items']['1_6']['memo'] == '推薦'

    assert revere_index['風味蛋炒飯'] == '0_0'
    assert revere_index['XO醬雙蛋(推薦)'] == '1_6'

def test_parse_taste():
    excel = openpyxl.load_workbook('商家資料填寫表格.xlsx')
    taste_sheet = excel['口味']
    reverse_idx = {'咖哩雞腿咖哩飯': '0_0', '香辣菜埔蛋炒飯': '0_1', '雞排麻婆豆腐飯': '1_0'}

    taste_map = main.parse_taste(taste_sheet, reverse_idx)

    assert taste_map['0_0'] == {'0_0_0': {'tasteName': '原味', 'price':0},
                                '0_0_1': {'tasteName': '辣味', 'price':0},
                                '0_0_2': {'tasteName': '加飯', 'price':10}}
    assert taste_map['0_1'] == {'0_1_0': {'tasteName': '加飯', 'price':10}}
    assert taste_map['1_0'] == {'1_0_0': {'tasteName': '加飯', 'price':10}}