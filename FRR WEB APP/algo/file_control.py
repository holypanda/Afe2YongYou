import os

def delete_direction(folder):
    walk = os.walk("%s/%s" % (os.getcwd(), folder))
    for root, dirs, files in walk:
        for file in files:
            path = "%s\%s" % (root, file)
            os.remove(path)