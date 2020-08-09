'0, 1, 1, 2, 3, 5, 8, 13, 21, 34

msgbox "Fibonacci Series"

i = 0
j = 1

strmsg = i& ", " &j& ", "

for x = 0 to 7

k = i + j

strmsg = strmsg &k& ", "

i = j
j = k

next

msgbox " The frist 10 fibonacci numbers are " &strmsg
