a = int(input())
b = int(input())
c = int(input())
pos = 0
neg = 0
for x in (a, b, c):
    if x > 0:
        pos += 1
    elif x < 0:
        neg += 1
print(pos)
print(neg)
