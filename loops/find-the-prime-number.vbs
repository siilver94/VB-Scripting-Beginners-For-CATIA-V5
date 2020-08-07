
msgbox " Welcome to the program to determine whether a number is prime or not"

dim prime, n

n = cint(inputbox("Enter the value for which you want to check"))

if n < 2 then
prime = false

elseif n = 2 then
prime = ture

else
	for i = 2 to (n-1)

		if (n mod i) = 0 then
		prime = flase
		exit for
		else
		prime = true
		end if
	next
end if

if prime = true then
msgbox " The number is prime"
else
msgbox " The number is not prime"
end if
