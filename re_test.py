a = []
a.append((1,2))
a.append((3,4))
print(a)

b,c = a.pop()

print(b,c)

import os

parent_path = os.path.dirname(os.path.realpath(__file__))
print(parent_path)