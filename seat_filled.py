#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 12 11:01:04 2020

@author: DAMASHII
"""
import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
from openpyxl.styles import PatternFill
import pandas as pd
import math

print('Seat Step: ', end='')
s_seat_step = int(input())

catwalk_size = 4
reserved_front = 5

if s_seat_step > catwalk_size:
    print('Must increase catwalk size')
    exit()

last_people_loc = []
color_i = 0
color_count = 0
reserveFill = PatternFill(fgColor='00FF00', fill_type='solid')
config_wb = openpyxl.load_workbook(filename='Config.xlsx')
list_wb = openpyxl.load_workbook(filename='List.xlsx')
row_no = 1
people_size = 1000


def main():
    block_info = pd.read_excel('Config.xlsx', sheet_name='Block Info')
    temp_inside = config_wb['Template Inside_2']
    print(block_info, '\n')

    config_wb.remove_sheet(config_wb['Template Filled Morning'])
    temp_filled_morning = config_wb.copy_worksheet(temp_inside)
    temp_filled_morning.title = 'Template Filled Morning'

    list_wb.remove_sheet(list_wb['list1'])
    list_wb.create_sheet('list1')
    list_1 = list_wb['list1']
    list_1.cell(row=1, column=1).value = 'No.'
    list_1.cell(row=1, column=2).value = 'Block'
    list_1.cell(row=1, column=3).value = 'Line'
    list_1.cell(row=1, column=4).value = 'Seat'

    list_wb.remove_sheet(list_wb['list2'])
    list_2 = list_wb.copy_worksheet(list_1)
    list_2.title = 'list2'

    blocks_seat_size = []
    for i in range(10):
        blocks_seat_size.append(
            count_available_block(block_info, i, temp_filled_morning))
    print('Total available chairs are\t', sum(blocks_seat_size))
    fill_special_block(block_info, blocks_seat_size, temp_filled_morning)

    musical_chair(list_2, list_1, block_info)

    config_wb.save('Config.xlsx')
    list_wb.save('List.xlsx')


def count_available_block(info, index, template):
    block_seat_count = 0
    block = info.at[index, 'Block']
    line_size = info.at[index, 'Line']
    seat_size = info.at[index, 'Seat']
    MAX_seat = info.at[index, 'Max Seat']
    side = info.at[index, 'Side']
    pivot = info.at[index, 'Pivot']

    beg_loc = coordinate_from_string(pivot)
    beg_col = column_index_from_string(beg_loc[0])
    beg_row = beg_loc[1]

    if side == 'L':
        line_step = -s_seat_step
        seat_step = -1
        print('Block', block, 'is Left Side')
    elif side == 'R':
        line_step = -s_seat_step
        seat_step = 1
        print('Block', block, 'is Right Side')
    elif side == 'C':
        line_step = 1
        seat_step = -s_seat_step
        print('Block', block, 'is Center Side')
    elif side == 'S':
        line_step = -2
        seat_step = s_seat_step
        seat_size = seat_size + catwalk_size
        print('Block', block, 'is Special Side')
    else:
        print('Block', block, 'is N/A Side')
        return None

    if side == 'L' or side == 'R':
        line_size, seat_size = seat_size, line_size

    print('Start\tat:', get_column_letter(beg_col), beg_row)
    end_col = beg_col+((seat_size-1)*(seat_step/abs(seat_step)))
    end_row = beg_row+((line_size-1)*(line_step/abs(line_step)))
    print('End\tat:', get_column_letter(end_col), end_row)

    if side == 'L' or side == 'R':
        line_size, seat_size = seat_size, line_size

    for i in range(line_size):
        seat_count = 0
        for j in range(int(seat_size/s_seat_step) + 1):
            # rotation block
            if side == 'C' or side == 'S':
                cur_line, cur_seat = i, j
            else:
                cur_line, cur_seat = j, i

            cur_row = beg_row+(cur_line*line_step)
            cur_col = beg_col+(cur_seat*seat_step)
            if side == 'S' and (cur_seat*seat_step) > (seat_size/2) - (catwalk_size/2):
                cur_col = cur_col + seat_step - 1

            if template.cell(row=cur_row, column=cur_col).value == 'x':
                if side == 'S':
                    template.cell(row=cur_row, column=cur_col).value = 'o'
                # template.cell(row=cur_row, column=cur_col).fill = reserveFill
                seat_count = seat_count + 1
                block_seat_count = block_seat_count + 1

        print('Line', block, i+1, '\thas', seat_count, '\tavailable chairs')
    if block_seat_count > MAX_seat:
        print('OVERFLOW Seat')

    print('available chairs are', block_seat_count, '\n')
    return block_seat_count


def fill_special_block(info, blocks_seat_size, template):
    s_loc = len(blocks_seat_size)-1

    # Import people
    p_info = import_people(blocks_seat_size)
    global people_size
    special_block_size = blocks_seat_size.pop(s_loc)
    remain_people_size = people_size - special_block_size

    info_copy = info
    if remain_people_size > 0:
        info = info.drop(s_loc)
        fill_upper_block(info, blocks_seat_size, template,
                         remain_people_size, p_info)
        global config_wb
        config_wb.save('Config.xlsx')
        print('Please check seatable chair are \'o\' sign in Excel')
        print('continue [Y/n]?', end='')
        resp = input()
        while(resp.upper() != 'Y'):
            if resp.upper() == 'N':
                exit()
            print('Please response only Y or N')
        config_wb = openpyxl.load_workbook(filename='Config.xlsx')
        template = config_wb['Template Filled Morning']
        reorder_upper(info, blocks_seat_size, template, p_info)

        fill_block(info_copy, s_loc, template, special_block_size, p_info, 'o')
    else:
        fill_block(info_copy, s_loc, template, people_size, p_info, 'o')


def fill_upper_block(info, blocks_seat_size, template, people_size, p_info):
    print('Remaining to upper people are\t', people_size)

    mid_loc = int(len(blocks_seat_size)/2)
    mid_seat_size = blocks_seat_size[mid_loc]
    remain_people_size = people_size - mid_seat_size

    if remain_people_size < 0:
        fill_block(info, mid_loc, template, people_size, p_info, 'x')
        return

    fill_block(info, mid_loc, template, mid_seat_size, p_info, 'x')
    for block_loc in range(mid_loc):
        left_loc = mid_loc-block_loc-1
        right_loc = mid_loc+block_loc+1

        left_seat_size = blocks_seat_size[left_loc]
        right_seat_size = blocks_seat_size[right_loc]

        if remain_people_size < left_seat_size+right_seat_size:
            if remain_people_size % 2:
                # odd number of people
                fill_block(info, left_loc, template,
                           int(remain_people_size/2)+1, p_info, 'x')
            else:
                # even number of people
                fill_block(info, left_loc, template, int(
                    remain_people_size/2), p_info, 'x')

            fill_block(info, right_loc, template,
                       int(remain_people_size/2), p_info, 'x')
            return

        remain_people_size = remain_people_size - left_seat_size - right_seat_size
        fill_block(info, left_loc, template, left_seat_size, p_info, 'x')
        fill_block(info, right_loc, template, right_seat_size, p_info, 'x')


def fill_block(info, index, template, people_size, p_info, sign):

    block_seat_count = 0
    block = info.at[index, 'Block']
    line_size = info.at[index, 'Line']
    seat_size = info.at[index, 'Seat']
    side = info.at[index, 'Side']
    pivot = info.at[index, 'Pivot']

    beg_loc = coordinate_from_string(pivot)
    beg_col = column_index_from_string(beg_loc[0])
    beg_row = beg_loc[1]

    if side == 'L':
        line_step = -s_seat_step
        seat_step = -1
    elif side == 'R':
        line_step = -s_seat_step
        seat_step = 1
    elif side == 'C':
        line_step = 1
        seat_step = -s_seat_step
    elif side == 'S':
        line_step = -2
        seat_step = s_seat_step
        seat_size = seat_size + catwalk_size
    else:
        print('Block', block, 'is N/A Side')
        return None

    for i in range(line_size):
        for j in range(int(seat_size/s_seat_step) + 1):
            # rotation block
            if side == 'C' or side == 'S':
                cur_line, cur_seat = i, j
            else:
                cur_line, cur_seat = j, i

            cur_row = beg_row+(cur_line*line_step)
            cur_col = beg_col+(cur_seat*seat_step)
            if side == 'S' and (cur_seat*seat_step) > (seat_size/2) - (catwalk_size/2):
                cur_col = cur_col + seat_step - 1

            if template.cell(row=cur_row, column=cur_col).value == sign:
                global color_i, color_count
                template.cell(row=cur_row, column=cur_col).value = 'o'
                template.cell(row=cur_row, column=cur_col).fill = PatternFill(
                    fgColor=p_info.at[color_i, 'C_code'], fill_type='solid')
                block_seat_count = block_seat_count + 1
                color_count = color_count + 1

                if sign == 'o':
                    global row_no
                    row_no = row_no + 1
                    list_1 = list_wb['list1']
                    list_1.cell(row=row_no, column=1).value = row_no - 1
                    list_1.cell(row=row_no, column=2).value = block
                    if side == 'S':
                        list_1.cell(row=row_no, column=3).value = i + 2
                        if (cur_seat*seat_step) > (seat_size/2) - (catwalk_size/2):
                            list_1.cell(row=row_no, column=4).value = (
                                (j+1)*s_seat_step) - catwalk_size
                        else:
                            list_1.cell(row=row_no, column=4).value = (
                                j*s_seat_step) + 1
                    else:
                        list_1.cell(row=row_no, column=3).value = i + 1
                        list_1.cell(row=row_no, column=4).value = (
                            j*s_seat_step) + 1

                if color_count == p_info.at[color_i, 'Size']:
                    color_i = color_i + 1
                    color_count = 0

                if block_seat_count == people_size:
                    last_people_loc.append([cur_row, cur_col])
                    print('Block', block, 'End', get_column_letter(
                        cur_col), cur_row, 'get people\t', people_size)
                    return


def import_people(blocks_seat_size):
    p_info = pd.read_excel('Order.xlsx', sheet_name='Morning Order')
    global people_size
    people_size = sum(p_info['Size'])
    print('Total people are\t\t', people_size)
    if people_size > sum(blocks_seat_size):
        print('Number of total people are overflow')
    return p_info


def reorder_upper(info, blocks_seat_size, template, p_info):
    global color_i, color_count
    color_i, color_count = 0, 0

    for block_loc in range(len(info)-1, -1, -1):
        seat_size = blocks_seat_size[block_loc]
        fill_block(info, block_loc, template, seat_size, p_info, 'o')


def musical_chair(list_out, list_in, info):
    global people_size, row_no

    # seat_size = info.at[len(info)-1, 'Seat']
    # last_reserved = int(reserved_front*seat_size/s_seat_step)
    first_reserved = 10
    last_reserved = 10
    rotate_size = row_no - first_reserved - last_reserved - 1

    print('Total Loop:', math.ceil((people_size-last_reserved) / rotate_size))

    for i in range(first_reserved):
        src_idx = i
        des_idx = i
        list_out.cell(row=des_idx+2, column=1).value = des_idx + 1
        for j in range(3):
            list_out.cell(row=des_idx+2, column=j+2).value = list_in.cell(row=src_idx+2, column=j+2).value

    for i in range(people_size-last_reserved):
        remainder = (people_size-first_reserved-last_reserved) % rotate_size
        src_idx = first_reserved + ((rotate_size-remainder+i) % rotate_size)
        list_out.cell(row=i+2, column=1).value = i + 1
        for j in range(3):
            list_out.cell(
                row=i+2, column=j+2).value = list_in.cell(row=src_idx+2, column=j+2).value

    for i in range(last_reserved):
        src_idx = rotate_size + i
        des_idx = people_size - last_reserved + i
        list_out.cell(row=des_idx+2, column=1).value = des_idx + 1
        for j in range(3):
            list_out.cell(row=des_idx+2, column=j+2).value = list_in.cell(row=src_idx+2, column=j+2).value


main()
