# Tax Sale and Foreclosure Parser

## What are foreclosures?

Foreclosures are more common knowledge to the layman than tax sales. In a foreclosure, when a mortgagor cannot pay their mortgage payments, the bank can put up the property for sale at a discounted price. Currently, foreclosures are announced in one large string of text.

## What are tax sales?

Tax sales are an essential concept in the housing industry. Properties whose property taxes are not paid off are auctioned in a monthly tax sale,
akin to how bank foreclose houses based on mortgages not being paid off. Local counties post the information about tax sales on their website or in the 
local newspaper. Currently, counties in Georgia post all information sent out in one large string of text.

## What does the project solve?

In order to make viewing the hundreds of properties embedded within these large information dumps easier without loss of information, this program allows for the user to define any fields they need. It can parse out each property's important characteristics. The program gives ultimate flexibility.

## How does it work?

As an example, a user might need every property's current owner's name, property's address, loan date, loan amount, and attorney. To find each field, each field has potential beginning parameters and potential endng parameters to search for. These arguments are inputted in an external options excel file. An example of an options file is below:

| Field Name | Beg ID  | End ID  |
| :---:   | :-: | :-: |
| Owner Name | Security Deed executed by <br /> Security Deed given by <br /> certain Security Deed from<br />Security Deed given from<br />to secure debt from <br />to secure debt given by | to <br />, to |
| Address | known as <br /> known as: | . <br />, together <br /> according to <br /> To <br /> Being |
| Loan Date | dated <br /> recorded on | , record <br /> and <br /> , and recorded <br /> . <br /> ( <br /> in <br /> at |
| Loan Amount | principal amount of <br /> in the amount of | ), <br /> , <br /> dated <br /> dollars, <br /> with |
| Attorney | Attorney in Fact for <br /> attorney-in-fact for | , |

Given this option file and source word document with the tax sale or foreclosure information, run the respective parser file (ForeclosureParser.py or TaxSaleParser.py) and a result excel sheet will be generated. An example result will look like this: 

![Example Result](./Images/Example%20Result.JPG?raw=true "Example Result")

Once outputted, format as needed, and you have all the information you need to quickly review all of a properties

## How to install

Developed on Python 3.7.9, so any version >= 3.7.9 is recommended.

Once cloned, run ``` pip install -r requirements.txt ``` to install the necessary packages. Populate the Source folder with a source.docx and options.xlsx with the format specified above. Then run ```python ForeclosureParser.py``` or ```python TaxSaleParser.py``` to run.

