let inputItemCount = 1;
grandTotal = 0;
const lineBreak = document.createElement('br');

//Buttons
const addItem = document.getElementById('addItem');
const submitBilling = document.getElementById('submitBilling');

document.addEventListener("DOMContentLoaded", function() 
{
    addItem.addEventListener('click', moreItem);
    submitBilling.addEventListener('click', generateBilling);
});

function moreItem() 
{
    inputItemCount++;

    const itemSection = document.getElementById('itemSection');

    // Create a new div for the item
    const newItemDiv = document.createElement('div');

    newItemDiv.innerHTML =  `
    <h3>Item ${inputItemCount}</h3>

    <label for="itemDesc${inputItemCount}" class="text">Item: </label>
    <input type="text" id="itemDesc${inputItemCount}" class="inputField"><br><br>

    <label for="itemQty${inputItemCount}" class="text">Quantity: </label>
    <input type="number" id="itemQty${inputItemCount}" class="inputField"><br>

    <label for="itemPrice${inputItemCount}" class="text">Price: </label>
    <input type="number" id="itemPrice${inputItemCount}" class="inputField"><br>

    <label for="currency${inputItemCount}" class="text">Currency: </label>
    <select id="currency${inputItemCount}" class="inputField">
      <option value="USD">USD</option>
      <option value="PHP">PHP</option>
    </select><br>
    `
    // Append the new item div to the container
    itemSection.appendChild(newItemDiv);
}

function convertCurrency()
{
    const getInputCurrency = [];
    for (let i = 1; i <= inputItemCount; i++) {
        const itemCurrency = document.getElementById('currency' + i).value;
        const itemPrice = document.getElementById('itemPrice' + i).value;

        // Item is in USD convert to PHP, vise versa
        if (itemCurrency === "USD") {
            const convertedPrice = itemPrice * 55.5; // Example conversion rate
            document.getElementById('itemPrice' + i).value = convertedPrice.toFixed(2);
        } else if (itemCurrency === "PHP") {
            const convertedPrice = itemPrice / 55.5; // Example conversion rate
            document.getElementById('itemPrice' + i).value = convertedPrice.toFixed(2);
        }
        getInputCurrency.push(itemCurrency);
    }
}

function formatCurrency(value, currency)
{
    if (currency === "USD") 
    {
        const formattedValue = new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: 2
        }).format(value);
    return formattedValue;
    }
    else if (currency === "PHP") 
    {
        const formattedValue = new Intl.NumberFormat('en-PH', {
        style: 'currency',
        currency: 'PHP',
        minimumFractionDigits: 2
        }).format(value);
    return formattedValue;
    }
}

function generateBilling() 
{
    //Get values from input fields
    const customerName = document.getElementById('customerName').value;
    const customerAttention = document.getElementById('customerAttention').value;
    const billingNum = document.getElementById('billingNum').value;
    const billingDate = document.getElementById('billingDate').value;
    const tinNum = document.getElementById('tinNum').value;
    const payDeadline = document.getElementById('payDeadline').value;
    const payMethod = document.getElementById('payMethod').value;

    fileName = billingNum + "Billing_" + customerName + ".docx";

    //convertCurrency();

    //Get item details
    const itemDetails = [];
    for (let i = 1; i <= inputItemCount; i++) {
        const itemNum = i;
        const itemDesc = document.getElementById('itemDesc' + i).value;
        const itemQty = document.getElementById('itemQty' + i).value;
        let itemPrice = document.getElementById('itemPrice' + i).value;

        // Get the selected currency
        const currency = document.getElementById('currency' + i).value;

        // Get subtotal of item * quantity
        let subtotal = itemPrice * itemQty;

        // Get grand total
        grandTotal += subtotal;

        // Add currency symbol to the price based on the selected currency
        itemPrice = formatCurrency(itemPrice, currency);
        subtotal = formatCurrency(subtotal, currency);

        itemDetails.push({ itemNum, itemDesc, itemPrice, itemQty, subtotal });
    }

    //Get Bank Details
    const bankDesc = document.getElementById('bankDesc').value;

    console.log("Generating voucher...");
    

    // Create a new Word document
    const 
    {
        Document,
        Paragraph,
        TextRun,
        Table,
        TableRow,
        TableCell,
        WidthType,
        AlignmentType,
    } = window.docx;
    
    
    // Helpers
    const heading = (text) =>
        new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new TextRun({
                    text: text,
                    bold: true,
                    font: "Montserrat",
                    size: 32,
                }),
            ]
        });
    
    const labelValue = (label, value = "") =>
        new Paragraph({
            children: [
            new TextRun({ 
                text: label + ": ", 
                bold: true,
                font: "Montserrat",
                size: 24
                }),
            new TextRun({
                text: value,
                font: "Montserrat",
                size: 24}),
          ],
        });
    
    const makeTable = (rows) =>
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: rows.map(
            (row) =>
              new TableRow({
                children: row.map(
                  (cell) =>
                    new TableCell({
                      width: { size: 20, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                                new TextRun({
                                text: cell,
                                AlignmentType: AlignmentType.CENTER,
                                font: "Montserrat",
                                size: 24,
                                }),
                            ],
                        })],
                    })
                ),
              })
          ),
        });

        const makeTotalTable = (rows) =>
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: rows.map(
                (row) =>
                  new TableRow({
                    children: row.map(
                      (cell) =>
                        new TableCell({
                          width: { size: 20, type: WidthType.PERCENTAGE },
                          children: [
                            new Paragraph({
                                alignment: AlignmentType.RIGHT,
                                children: [
                                    new TextRun({
                                    text: cell,
                                    AlignmentType: AlignmentType.LEFT,
                                    font: "Montserrat",
                                    size: 24,
                                    }),
                                ],
                            })],
                        })
                    ),
                  })
              ),
            });
    
    const makeParagraph = (text) =>
        new Paragraph({
            children: [
                new TextRun({
                    text: text,
                    font: "Montserrat",
                    size: 24,
                    indent: {
                        left: 720,
                        hanging: 360,
                    },
                }),
            ],
        });

    // Create a table for grand total
    const grandTotalCurrency = document.getElementById('currency1').value;
    const grandTotalRows = [
        ["Grand Total", formatCurrency(grandTotal, grandTotalCurrency)],
    ];

    // Create a table for item details
        const itemRows = [
            ["Item", "Description", "Price", "Quantity", "Subtotal"]
        ];

    // Add data rows from iitem Details
    itemDetails.forEach(item => 
    {
        //Turn day to string
        const itemNumber = item.itemNum;
        //Convert to string
        item.itemNum = itemNumber.toString();
        itemRows.push([
            item.itemNum,
            item.itemDesc,
            item.itemPrice,
            item.itemQty,
            item.subtotal   
            ]);
    });
    
    const doc = new Document({
        sections: [
        {
            properties: 
            {
                page: 
                {
                    size: 
                    {
                        width: 11906,
                        height: 16838,
                    }
                }
            },
            children: [
            heading("STATEMENT OF ACCOUNT"),
                
            labelValue("Bill No. ", billingNum),
            labelValue("Bill To: ", customerName),
            labelValue("Attention To ", customerAttention),
            labelValue("Date ", billingDate),
            labelValue("TIN ", tinNum),
            labelValue("Deadline of Payment ", payDeadline),
            labelValue("Payment Method ", payMethod),

            //Item Details
            heading("Item Details"),
            makeTable(itemRows),
            makeTotalTable(grandTotalRows),
            
            //Bank Details
            heading("Bank Details"),
            makeParagraph(bankDesc),

            heading("Notes"),
            makeParagraph("All Black Markings are to be paid. All Red Markings are paid."),
            ],
        },
        ],
    });
    docx.Packer.toBlob(doc).then((blob) => 
    {
        console.log(blob);
        saveAs(blob, fileName);
        console.log("Document created successfully");
    });
}
  