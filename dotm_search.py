#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "SmileySlays"

import os
import subprocess
import argparse


def doc_to_text_catdoc(filename):
    PIPE = subprocess.PIPE
    p = subprocess.Popen(
        ' '.join(['textutil', '-stdout', '-cat', 'txt', filename]), shell=True, stdin=PIPE, stdout=PIPE, stderr=PIPE, close_fds=True)
    (child_stdin, child_stdout, child_stderr) = (p.stdin, p.stdout, p.stderr)

    return child_stdout.read()


def main():

    count = 0
    dir_ = "/Users/user/Documents/KAQ3/backend-dotm-search-assessment/dotm_files"
    if args.dir:
        dir_ = "/Users/user/Documents/KAQ3/backend-dotm-search-assessment" + \
            args.dir[1:]
    os.chdir(dir_)
    for files in os.listdir():
        if files.endswith(".dotm"):
            temp = doc_to_text_catdoc(files).decode("utf-8")
            if args.search in temp:
                count += 1
                print("Match found in file " + files)
                for index, character in enumerate(temp):

                    if character == args.search:
                        search_phrase = temp[max(
                            0, index-40):min(len(temp), index+40)]
                        # for num in range(1, 81):
                        #     try:
                        #         count += temp[index - 40 + num]
                        #     except:
                        #         continue
                        print(search_phrase)
                    else:
                        continue
    print("Total times '$' is found:" + str(count))


parser = argparse.ArgumentParser(
    description='Return instance of possible bills')
parser.add_argument('search', type=str,
                    help='What to search for')
parser.add_argument('--dir', type=str,
                    help='Directory to search ')
args = parser.parse_args()
if __name__ == '__main__':
    main()
