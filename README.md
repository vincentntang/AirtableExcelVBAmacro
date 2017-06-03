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

See blog post for more details (later will update)

