msgbox "Start"

vf = 1

vsum = cint(inputbox("Enter the value to find its factorial"))

for i = 1 to vsum step 1

	vf = vf * i

	if vsum > 15 then

	exit for

	end if

	next 

		if vsum > 15 then
			msgbox "Enter a value lesser then 15"
			else 
				msgbox " The factorial value i s" &vf

		end if
