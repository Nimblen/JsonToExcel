import argparse


parser = argparse.ArgumentParser()


parser.add_argument('--input', '-i', dest="input", type=str, help='Путь к входному файлу JSON')
#parser.add_argument('--excel', '-ex', dest="excel", help='Входные данные в формате Excel', default=False, action='store_true')
parser.add_argument('--output', '-o', dest="output", type=str, help='Путь к выходному файлу Excel')
parser.add_argument('--sheet-name', '-s', type=str, help='Имя листа в файле Excel', default='Sheet1')

