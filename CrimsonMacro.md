# Crimson RAT Sample

[Download sample](https://app.any.run/tasks/206fb61a-38ac-4f84-81f6-9389ce775c16/) 

# Document Password Removal

The document's macros are password protected but bypassing these is quite easy. The method I've used was posted on [Stackoverflow](https://stackoverflow.com/questions/272503/removing-the-password-from-a-vba-project)

Simply open the document in a text editor and search for the string 'DPB' & replace it with the string 'DPx'.

![](/pic-set3/img1.png)

After replacing those strings, open up the document and go to the macros tab (Developer > Visual Basic). You will likely get this popup:

![](/pic-set3/img3.png)

After clicking yes, you may get this error popup as well, simply ignore it.

![](/pic-set3/img4.png)

Once in the Visual Basic window, go to Tools > Project Properties > Protection & set a new password then exit out of the document & save the changes.

Now when you go into the Developer > Visual Basic tab and look at the macros, you should be able to view them by inputting the new password

# Macro Analysis

This first subroutine (named **Con**) sets a Path variable to: `C:\Users\Username\Intel.exe` which tells us that this macro will likely be dropping a PE file named **Intel.exe**

![](/pic-set3/img5.png)

The **Con** Subroutine also splits the data in **UserForm1** by the value of an exclamation point, all this means it simply splits the data by an exclamation point. This tells us that there's data in the **UserForm1** form.

Viewing that form, we don't see anything at first glance:

![](/pic-set3/img6.png)

But expanding the form, we see a small box & when we expand it, we see data:

![](/pic-set3/img7.png)

![](/pic-set3/img8.png)

So the **Con** subroutine will be splitting this data by exclamation points meaning it will simply be removing the exclamation points

![](/pic-set3/img5.png)

I am not too sure of the exact process this goes through to turn the data from the numbers into bytes but this is my best interpretation:

1. It removes the exclamation point from the data & turns it into an array
2. It loops through each element in the array 
3. It turns each array element (each set of numbers) into a Byte 

Moving on to the next subroutine:

![](/pic-set3/img9.png)

This one is pretty simple, it just adds the value `C:\Users\Username\Intel.exe` (where Username is the username of the victim) to the Path1 string & opens & writes the bytes from the last subroutine to the **Intel.exe** file

And onto the last subroutine:

![](/pic-set3/img10.png)

As you could've guessed by the name, this subroutine simply runs the **Intel.exe** file with a window focus of 1 which means that the Window will be seen by the victim

Testing our theory that the window should be seen by the victim:

![](/pic-set3/img13.png)

# Running the Macros

After running the macros, I didn't see the window popup but I assume that the **Intel.exe** file had some code that hid its window.

But we can see that it is indeed ran:

![](/pic-set3/img11.png)

The **Intel.exe** file reaches out to **51.89.208.53:350**, it also tries to reach the same host on port 8730:

![](/pic-set3/img12.png)

Attempting to view the communication with Wireshark didn't yield much:

![](/pic-set3/img13.png)

# Links Used

1. [Converting ASCII to hex](https://www.asciitohex.com/)

2. [VBA Functions List](https://www.excelfunctions.net/vba-functions.html)

3. [MSDN Shell()](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shell-function)