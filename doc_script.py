#!/usr/bin/python3

import markdown

def create_docs(filename):
    with open(filename, 'r') as md_file:
        data = md_file.read()

    text = markdown.markdown(data)

    with open('index.html', 'w') as idx_file:
        idx_file.write(text)
    return

if __name__ == '__main__':
    filename = 'README.md'
    create_docs(filename)
