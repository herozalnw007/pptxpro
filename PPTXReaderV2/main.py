'''
Created on Jan 22, 2019

@author: Bowonwit
'''
from __future__ import print_function, unicode_literals
from pptx import Presentation
from PyInquirer import prompt, print_json
import os
import time

import extract_pptx as e_p
from extract_pptx import *

def main():

    while True:
        default_path = 'C:/Users/Bowonwit/Desktop/'
        _path = ''
        _path = default_path
        aws_path = False
        
        while True:
            files_names = ['Root...']
            que_path = [
                {
                    'type': 'input',
                    'name': 'path',
                    'message': 'Please Input the path of your PPTX file that you will execute.(\"Exit\" to close program)',
                    'default': _path,
                },
            ]
            aws_path = prompt(que_path)
            if aws_path['path'].lower() == 'exit':
                print ("End,", end='')
                for i in range(28):
                    time.sleep(0.05)
                    print('.', end=' ')
                print ("\nGood bye !!!")
                exit()
            else:
                if aws_path['path'][-1] != '/': aws_path['path'] += '/'
                for root_list, dirs_list, files_list in os.walk(aws_path['path']):
                    for _file in files_list:
                        if str(_file)[-5:] == ".pptx" and str(_file)[0:2] != "~$":
                            files_names.append(str(_file))
                if len(files_names) <= 1:
                    print ("The program not found PPTX file in this path. Please try again!", end=' ')
                    for i in range(5):
                        time.sleep(0.01)
                        print('.', end=' ')
                    print('.',)
                else:
                    while True:
                        que_file = [
                            {
                                'type': 'list',
                                'name': 'file',
                                'message': 'Which one is your PPTX file that you will execute?',
                                'choices': files_names,
                            },
                        ]
                        aws_file = prompt(que_file)
                        if aws_file['file'] == 'Root...': break
                        else:
                            
                            while True:
                                que_function = [
                                    {
                                        'type': 'list',
                                        'name': 'functions',
                                        'message': 'Please selcet a function that you want to use.',
                                        'choices': ['  1). Extract PPTX file.', '  2). Get **********(Unavaliable).','  3). Get *********(Unavaliable).','  4). Get Analysis.',' Back...'],
                                    }
                                ]
                                aws_function = prompt(que_function)
                                
                                if aws_function['functions'] == que_function[0]['choices'][-1]: break
                                elif aws_function['functions'] == que_function[0]['choices'][0]: e_p.handle_extract(aws_path['path'], aws_file['file'])
                                elif aws_function['functions'] == que_function[0]['choices'][1]: print ("This function's Unavaliable yet.111")
                                elif aws_function['functions'] == que_function[0]['choices'][2]: print ("This function's Unavaliable yet.222")
                                elif aws_function['functions'] == que_function[0]['choices'][3]: print ("This function's Unavaliable yet.333")


if __name__ == '__main__':
    main()