# What does this macro do?

This is an excel VBA macro that does the following:

- Takes in two columns of data (PrimaryKey, Airtable Download attachment)
- Processes "attachment" field so both the image name + image URL are in their own seperate lines
- Creates a copy-ready batch command that you can put in a .bat file to creating all your image asset names

It outputs the following columns of data

- Column 1 = primary key
- Column 2 = airtable's URL image link
- Column 3 = airtables original file name on upload
- Column 4 = batch commands for duplicating / creating images

Column 2 is used for the bulk image downloader of your choice (ExtremePictureFinder)

Column 4 is for windows .bat file used for duplicating / renaming files

# How do I use this macro?

Open up excel

Press ALT+F11

Create a module

Paste the code

Press F5

# Where can I read more about the entire process?

See blog post for more details (later will update)

# What is ultimate goal of this excel VBA?

Make all image assets locally from airtable. Ideally the least number of steps would be:

1. Download the CSV file (Col 1:PrimaryKey, Col2: ImageLink) say `C:/AirtableDL`
2. Click one button on the ribbon menu (Creates 4 columns of data, automatically downloads image col 2 in same location, automatically create and run batch file)
3. All assets finished sitting inside same directory

Would solve a lot of issues related to digitally managing image assets in 2 clicks

