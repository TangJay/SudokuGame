from venv import create
import numpy as np
import pandas as pd
import sys
import copy
import time
import random

def get_sudoku(file_name):
    df = pd.read_excel(file_name,header = None)
    sudoku = df.values
    print(sudoku)
    return sudoku
    

def check_clues(sudoku, x,y):
    clues = set()
    num = [1,2,3,4,5,6,7,8,9]
    for i in range(9):
        if sudoku[i][y] != 0:
            clues.add(sudoku[i][y])
    for i in range(9):
        if sudoku[x][i] != 0:
            clues.add(sudoku[x][i])
    b_x = x // 3
    b_y = y // 3
    for i in range(3):
        for j in range(3):
            if sudoku[b_x * 3 + j][b_y * 3 + i] != 0:
                clues.add(sudoku[b_x * 3 + j][b_y * 3 + i])    
    return set(num) - set(clues)
def find_zero(sudoku):
    for i in range(81):
        x = i % 9
        y = i // 9
        if sudoku[x][y] == 0:
            return x, y
    return -1, -1
def solve_sudoku(sudoku, times, ans):
    # 找到還沒填的那一格的x, y
    if len(ans) >= 1:
        return sudoku, times
    times += 1
    x, y = find_zero(sudoku)
    if x == -1 and y == -1:
        ans.append(sudoku)
    clues = check_clues(sudoku, x, y)
    
    if len(clues) == 0:
        times -= 1
        return 
    elif len(clues) == 0:
        return sudoku
    for m in clues:
        new_sudoku = copy.deepcopy(sudoku)
        new_sudoku[x][y] = m
        solve_sudoku(new_sudoku, times, ans)

def show_answers(ans):
        showAns = input("Do you want to show one answers?     (print yes or no)")
        if showAns == "yes":
            print(ans)
        elif showAns == "no":
            return
        else:
            print("that is not yes or no")
            show_answers(ans)

def create_sudoku(nums,times):
    numbers = [0,1,2,3,4,5,6,7,8]
    numbers_in_sudoku = list()
    sudoku = np.zeros((9,9))
    while True:
        present_nums = 0
        for i in range(9):
            for j in range(9):
                if sudoku[i,j] != 0:
                    present_nums += 1
        
        if present_nums == nums:
            break
        x = list()
        y = list()
        for i in range(nums-present_nums):
            x += random.sample(numbers, 1)
            y += random.sample(numbers, 1)
        
        
        for i in range(nums-present_nums):
            clues = list(check_clues(sudoku, x[i],y[i]))
            if len(clues) == 0:
                break
            random.shuffle(clues)
            sudoku[x[i], y[i]] = clues[0]
        
    ans = list()
    solve_sudoku(sudoku,0,ans)
    if len(ans) == 0:
        create_sudoku(30, times + 1)
    elif len(ans) >= 1:
        print("times = ", times)
        return sudoku
        print("True")
        print(sudoku)
        print("create", times, "times")
        show_answers(ans[0])
    
    
if __name__ == '__main__':
    #basicBox = generate_sudoku()
    
    sudoku = create_sudoku(30,1)
    
    

    #check_sudoku(basicBox)