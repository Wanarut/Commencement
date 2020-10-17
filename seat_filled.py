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

# load_people = 1900
print('Seat Step: ', end='')
s_seat_step = int(input())

catwalk_size = 4
if catwalk_size < s_seat_step:
    catwalk_size = s_seat_step

last_people_loc = []
reserveFill = PatternFill(fgColor='00FF00', fill_type='solid')


def main():
    workbook = openpyxl.load_workbook(filename='Config.xlsx')
    block_info = pd.read_excel('Config.xlsx', sheet_name='Block Info')
    temp_inside = workbook['Template Inside_2']
    print(block_info, '\n')

    workbook.remove_sheet(workbook['Template Filled'])
    temp_filled = workbook.copy_worksheet(temp_inside)
    temp_filled.title = 'Template Filled'

    blocks_seat_size = []
    for i in range(10):
        blocks_seat_size.append(
            count_available_block(block_info, i, temp_inside))
    print('Total available chairs are\t', sum(blocks_seat_size))
    fill_special_block(block_info, blocks_seat_size, temp_filled)

    workbook.save('Config.xlsx')


def count_available_block(info, index, template):
    block_seat_count = 0
    block = info.at[index, 'Block']
    line_size = info.at[index, 'Line']
    seat_size = info.at[index, 'Seat']
    MAX_seat = info.at[index, 'Max Seat']
    side = info.at[index, 'Side']
    pivot = info.at[index, 'Pivot']

    ref_loc = coordinate_from_string(pivot)
    ref_col = column_index_from_string(ref_loc[0])
    ref_row = ref_loc[1]

    if side == 'L':
        line_step = -1
        seat_step = -1
        print('Block', block, 'is Left Side')
    elif side == 'R':
        line_step = -1
        seat_step = 1
        print('Block', block, 'is Right Side')
    elif side == 'C':
        line_step = 1
        seat_step = -1
        print('Block', block, 'is Center Side')
    elif side == 'S':
        line_step = 2
        seat_step = s_seat_step
        seat_size = seat_size + catwalk_size
        print('Block', block, 'is Special Side')
    else:
        print('Block', block, 'is N/A Side')
        return None

    if side == 'L' or side == 'R':
        line_size, seat_size = seat_size, line_size

    print('Start\tat:', get_column_letter(ref_col), ref_row)
    print('End\tat:', get_column_letter(ref_col + ((seat_size-1)*seat_step)),
          ref_row + ((line_size-1)*line_step))

    if side == 'L' or side == 'R':
        line_size, seat_size = seat_size, line_size

    for i in range(line_size):
        seat_count = 0
        for j in range(seat_size):
            # rotation block
            if side == 'C' or side == 'S':
                cur_line, cur_seat = i, j
            else:
                cur_line, cur_seat = j, i

            if cur_seat*seat_step > seat_size:
                break

            cur_row = ref_row+(cur_line*line_step)
            if side == 'S' and cur_seat*seat_step >= (seat_size/2)-(catwalk_size/2):
                # offset = int(catwalk_size/2) + ((seat_step+1) % 2)
                cur_column = ref_col+(cur_seat*seat_step)+catwalk_size
            else:
                cur_column = ref_col+(cur_seat*seat_step)

            if template.cell(row=cur_row, column=cur_column).value == 'x':
                seat_count = seat_count + 1
                block_seat_count = block_seat_count + 1

        print('Line', block, i+1, '\thas', seat_count, '\tavailable chairs')
    if block_seat_count > MAX_seat:
        print('OVERFLOW Seat')

    print('available chairs are', block_seat_count, '\n')
    return block_seat_count


def fill_special_block(info, blocks_seat_size, template):
    s_loc = len(blocks_seat_size)-1

    block_seat_count = 0
    block = info.at[s_loc, 'Block']
    line_size = info.at[s_loc, 'Line']
    seat_size = info.at[s_loc, 'Seat']
    pivot = info.at[s_loc, 'Pivot']

    ref_loc = coordinate_from_string(pivot)
    ref_col = column_index_from_string(ref_loc[0])
    ref_row = ref_loc[1]

    line_step = 2
    seat_step = s_seat_step
    seat_size = seat_size + catwalk_size

    print('People Size: ', end='')
    load_people = int(input())
    people_size = load_people
    print('Total people are\t\t', people_size)
    if people_size > sum(blocks_seat_size):
        print('Number of total people are overflow')
        return

    print(blocks_seat_size)
    special_block_size = blocks_seat_size.pop(s_loc)
    remain_people_size = people_size - special_block_size

    for cur_line in range(line_size):
        for cur_seat in range(seat_size):
            if cur_seat*seat_step > seat_size:
                break

            cur_row = ref_row+(cur_line*line_step)
            # if cur_seat*seat_step >= seat_size/2:
            if cur_seat*seat_step >= (seat_size/2)-(catwalk_size/2):
                # offset = int(catwalk_size/2) + ((seat_step+1) % 2)
                # cur_column = ref_col+(cur_seat*seat_step)+offset
                cur_column = ref_col+(cur_seat*seat_step)+catwalk_size
            else:
                cur_column = ref_col+(cur_seat*seat_step)

            if template.cell(row=cur_row, column=cur_column).value == 'x':
                template.cell(
                    row=cur_row, column=cur_column).fill = reserveFill
                block_seat_count = block_seat_count + 1
                # print(block_seat_count)

                if block_seat_count == people_size or block_seat_count == special_block_size:
                    last_people_loc.append([cur_row, cur_column])
                    print('Block', block, 'End', get_column_letter(cur_column),
                          cur_row, 'get people\t', block_seat_count)

                    info = info.drop(s_loc)
                    if remain_people_size > 0:
                        fill_upper_block(info, blocks_seat_size,
                                         template, remain_people_size)
                    return


def fill_upper_block(info, blocks_seat_size, template, people_size):
    print('Remaining to upper people are\t', people_size)

    mid_loc = int(len(blocks_seat_size)/2)
    mid_seat_size = blocks_seat_size[mid_loc]
    remain_people_size = people_size - mid_seat_size

    if remain_people_size < 0:
        fill_block(info, mid_loc, template, people_size)
        return

    fill_block(info, mid_loc, template, mid_seat_size)
    for block_loc in range(mid_loc):
        left_loc = mid_loc-block_loc-1
        right_loc = mid_loc+block_loc+1

        left_seat_size = blocks_seat_size[left_loc]
        right_seat_size = blocks_seat_size[right_loc]

        if remain_people_size < left_seat_size+right_seat_size:
            if remain_people_size % 2:
                # มีแบ่งครึ่งมีเศษ 1 คน
                fill_block(info, left_loc, template,
                           int(remain_people_size/2)+1)
            else:
                # แบ่งเท่า
                fill_block(info, left_loc, template, int(remain_people_size/2))

            fill_block(info, right_loc, template, int(remain_people_size/2))
            return

        remain_people_size = remain_people_size - left_seat_size - right_seat_size
        fill_block(info, left_loc, template, left_seat_size)
        fill_block(info, right_loc, template, right_seat_size)


def fill_block(info, index, template, people_size):

    block_seat_count = 0
    block = info.at[index, 'Block']
    line_size = info.at[index, 'Line']
    seat_size = info.at[index, 'Seat']
    MAX_seat = info.at[index, 'Max Seat']
    side = info.at[index, 'Side']
    pivot = info.at[index, 'Pivot']

    ref_loc = coordinate_from_string(pivot)
    ref_col = column_index_from_string(ref_loc[0])
    ref_row = ref_loc[1]

    if side == 'L':
        line_step = -1
        seat_step = -1
        # print('Block', block, 'is Left Side')
    elif side == 'R':
        line_step = -1
        seat_step = 1
        # print('Block', block, 'is Right Side')
    elif side == 'C':
        line_step = 1
        seat_step = -1
        # print('Block', block, 'is Center Side')
    elif side == 'S':
        line_step = 2
        seat_step = 1
        seat_size = seat_size + catwalk_size
        # print('Block', block, 'is Special Side')
    else:
        print('Block', block, 'is N/A Side')
        return None

    if side == 'L' or side == 'R':
        line_size, seat_size = seat_size, line_size

    # print('Start\tat:', get_column_letter(ref_col), ref_row)
    # print('End\tat:', get_column_letter(ref_col + ((seat_size-1)*seat_step)),
    #       ref_row + ((line_size-1)*line_step))

    if side == 'L' or side == 'R':
        line_size, seat_size = seat_size, line_size

    for i in range(line_size):
        for j in range(seat_size):
            if side == 'C' or side == 'S':
                cur_line, cur_seat = i, j
            else:
                cur_line, cur_seat = j, i

            cur_row = ref_row+(cur_line*line_step)
            cur_column = ref_col+(cur_seat*seat_step)

            if template.cell(row=cur_row, column=cur_column).value == 'x':
                template.cell(
                    row=cur_row, column=cur_column).fill = reserveFill
                block_seat_count = block_seat_count + 1

                # if block_seat_count > people_size:
                #     return

                if block_seat_count == people_size:
                    last_people_loc.append([cur_row, cur_column])
                    print('Block', block, 'End', get_column_letter(
                        cur_column), cur_row, 'get people\t', people_size)
                    return


main()
