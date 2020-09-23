All you need to do to get this to work is..

1) Open your project
2) Hit Ctrl+T
3) Goto "Browse" and select my RazorDialog.OCX as a control (wherever you unzipped it)
4) Then take the control and draw it on your form
5) Make sure you have a RichTextBox control, it will not work with anything else



Then lets say you wanted to open a file into the RichTextBox you have..
You would do something like this:


RazorDialog1.OpenFileDefaultFilter RichTextBox1


Thats it! And you would put that code into a button, and once you clicked the button you could open a file! Remember, you can substitude the RazorDialog1 and RichTextBox1 names for whatever you, their just the default names.

===============================
The controls you can use
===============================

RazorDialog1.ShowAbout
RazorDialog1.SetFont RichTextBox
RazorDialog1.SetBackColor RichTextBox
RazorDialog1.SetTextColor RichTextBox
RazorDialog1.OpenFileDefaultFilter RichTextBox
RazorDialog1.OpenFileCustomFilter RichTextBox, Filter
RazorDialog1.SaveFileDefaultFilter RichTextBox
RazorDialog1.SaveFileCustomFilter RichTextBox, Filter