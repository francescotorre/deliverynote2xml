# DeliveryNote2XML

## Introduction

**DeliveryNote2XML** converts dynamically generated shipping documents from RTF format to XML.

## Warning

In order to function properly, **DeliveryNote2XML** could close **Microsoft Word/Excel** *without prompt*.

So, **make sure you have saved any open documents before launching the program**.

## How to Use This Program

First, the program is tailored to a specific document structure, so you need a delivery note in the correct format. For that, you can use one of the samples in the **samples** folder or you can fill in a template yourself.

Empty templates can be found in the **templates** folder.

Once you have a valid delivery note, follow these steps:

1. Launch **DeliveryNote2XML**

2. Select your delivery note by clicking **Open RTF file**

3. Optionally, click **Save XML file as. . .** to choose the location where the generated XML file will be saved. By default, it's the same folder as the RTF document

4. Click **CONVERT**

## How to Fill In a Delivery Note Template

There are fields in the template.

Each field contains a specific piece of information (e.g. company name, item code, item description, quantity, ecc.).

Here are some important points:

1. Fields in the upper left-hand corner of the document hold data about the seller (You).

2. Fields in the SHIP TO and SOLD TO sections hold data about the buyer (that is: your customer).

    **Note: DeliveryNote2XML looks in the customer data file for discounts to be applied to each customer.**
    This means the customer id in each delivery note must be found in the customer data file, otherwise a warning will be issued.

    You can change this file location by clicking on the **Change** button in the UI

3. Just below the SHIP TO and SOLD TO sections, you can find fields for:

    \- Order date

    \- Order number

    \- Document date

    \- Document number...

    ...and so on.

    most fields have a label that is self-explanatory

4. Next follows item data. Each line in this section identifies one product: code, description and quantity.

    Note the first line is a sample and it's pre-filled

5. Make sure to remove unused item fields, or the program will crash.

   **Note: empty non-item fields need not to be removed.**

6. More fields follow further on down in the template

## Unused Fields

Some fields are empty.

They can be used in the future to add more data, such as: 

\- vat code

\- start/end transport date/time

\- a footnote.

Click inside any field and use arrow keys to move to the previous/next field.

Obviously, the source code of the application would need to be modified to account for the new fields.

## Exception Handling

Sometimes documents can have an invalid structure.

A common glitch is when customer information is incorrectly placed in the DESCRIPTION field of the first item.

Look at the files in the **issues** folder to get an idea of how **DeliveryNote2XML** handles this occurrence.
