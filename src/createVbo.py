import os
import sys
import re
import time

def create_vbo(filename):
    full_path = filename + ".vbo"
    file = open(full_path, 'w', encoding="utf-8")
    new_str = "File created on " + str(time.strftime("%d/%m/%Y", time.localtime(
    ))) + " at " + str(time.strftime("%H:%M:%S", time.localtime()))
    file.write(new_str)
    file.write("\n")
    file.write("\n")
    file.write("[header]")
    file.write("\n")
    file.close()
    return full_path

if __name__ == '__main__':
    create_vbo(sys.argv[1])
