let inputFlightCount = 1;
let inputItineraryCount = 1;
const lineBreak = document.createElement('br');

//Buttons
const addFlight = document.getElementById('addFlight');
const addItinerary = document.getElementById('addItinerary');
const submitVoucher = document.getElementById('submitVoucher');

document.addEventListener("DOMContentLoaded", function() 
{
    addFlight.addEventListener('click', moreFlight);
    addItinerary.addEventListener('click', moreItinerary);
    submitVoucher.addEventListener('click', generateVoucher); 
});

function moreFlight() 
{
    inputFlightCount++;

    const flightSection = document.getElementById('flightSection');
    const newFlightContainer = document.createElement('div');

    newFlightContainer.innerHTML =  `
    <label for="flightNum${inputFlightCount}" class="text">Flight No.: </label>
    <input type="text" placeholder="Flight No." id="flightNum${inputFlightCount}" class="inputField">
    <br>
    <label for="departureDate${inputFlightCount}" class="text">Departure Date: </label>
    <input type="date" id="departureDate${inputFlightCount}" class="inputField">

    <label for="departureTime${inputFlightCount}" class="text">Departure Time: </label>
    <input type="time" id="departureTime${inputFlightCount}" class="inputField">
    <br>
    <label for="arrivalDate${inputFlightCount}" class="text">Arrival Date: </label>
    <input type="date" id="arrivalDate${inputFlightCount}" class="inputField">
    
    <label for="arrivalTime${inputFlightCount}" class="text">Arrival Time: </label>
    <input type="time" id="arrivalTime${inputFlightCount}" class="inputField">
    <br>
    <br>
    `
    // Append the new item div to the container
    flightSection.appendChild(newFlightContainer);
}

function moreItinerary()
{
    inputItineraryCount++;

    const itinerarySection = document.getElementById('itinerarySection');
    const newItineraryContainer = document.createElement('div');

    newItineraryContainer.innerHTML =  `
          <h3>Day ${inputItineraryCount}</h3>

      <label for="itiCity${inputItineraryCount}" class="text">City: </label>
      <input type="text" placeholder="City" id="itiCity${inputItineraryCount}" class="inputField">

      <label for="itiHotel${inputItineraryCount}" class="text">Hotel: </label>
      <input type="text" placeholder="Hotel" id="itiHotel${inputItineraryCount}" class="inputField">

      <label for="bf${inputItineraryCount}" class="text">Breakfast</label>
      <input type="checkbox" id="bf${inputItineraryCount}" class="checkbox">

      <label for="l${inputItineraryCount}" class="text">Lunch</label>
      <input type="checkbox" id="l${inputItineraryCount}" class="checkbox">

      <label for="d${inputItineraryCount}" class="text">Dinner</label>
      <input type="checkbox" id="d${inputItineraryCount}" class="checkbox">
      <br>
      <label for="desc${inputItineraryCount}" class="text">Description: </label>
      <textarea id="desc${inputItineraryCount}" rows="5" cols="50" placeholder="Description" class="inputField"></textarea>
      <br>
    `
    itinerarySection.appendChild(newItineraryContainer);
}

function generateVoucher() 
{
    //Separate per line in text area to array
    //Get values from input fields
    const passengerNames = document.getElementById('passengerNames').value;
    const passengerNamesArray = passengerNames.split('\n');
    console.log(passengerNamesArray);

    //Get Scope of Work
    const incDesc = document.getElementById('incDesc').value;
    const incDescArray = incDesc.split('\n');

    const excDesc = document.getElementById('excDesc').value;
    const excDescArray = excDesc.split('\n');

    const arrIns = document.getElementById('arrIns').value;
    const arrInsArray = arrIns.split('\n');

    const rem = document.getElementById('rem').value;
    const remArray = rem.split('\n');

    //Get values from input fields
    const customerName = document.getElementById('customerName').value;
    const voucherNum = document.getElementById('voucherNum').value;
    const voucherDate = document.getElementById('voucherDate').value;
    const destination = document.getElementById('destination').value;
    const travelDate = document.getElementById('travelDate').value;
    const hotelName = document.getElementById('hotelName').value;
    const hotelConf = document.getElementById('hotelConf').value;
    const hotelAddress = document.getElementById('hotelAddress').value;
    const rooms = document.getElementById('rooms').value;
    

    fileName = voucherNum + "Voucher" + "_" + customerName + ".docx";

    //Get flight details
    const flightDetails = [];
    for (let i = 1; i <= inputFlightCount; i++) {
        const flightNum = document.getElementById('flightNum' + i).value;
        const departureDate = document.getElementById('departureDate' + i).value;
        const departureTime = document.getElementById('departureTime' + i).value;
        const arrivalDate = document.getElementById('arrivalDate' + i).value;
        const arrivalTime = document.getElementById('arrivalTime' + i).value;
        flightDetails.push({ flightNum, departureDate, departureTime, arrivalDate, arrivalTime });
    }
    console.log(flightDetails);

    //Get Itinerary details
    const itineraryDetails = [];
    for (let i = 1; i <= inputItineraryCount; i++) {
        const day = i 
        const city = document.getElementById('itiCity' + i).value;
        const hotelName = document.getElementById('itiHotel' + i).value;
        const bf = document.getElementById('bf' + i).checked ? 'Breakfast' : '';
        const l = document.getElementById('l' + i).checked ? 'Lunch' : '';
        const d = document.getElementById('d' + i).checked ? 'Dinner' : '';

        //Get description
        const desc = document.getElementById('desc' + i).value;
        const descArray = desc.split('\n');

        itineraryDetails.push({day, city, hotelName, bf, l, d, descArray });
    }
    console.log(itineraryDetails);

    

    console.log("Generating voucher...");
    

    // Create a new Word document
    const 
    {
        Document,
        Paragraph,
        PageOrientation,
        PageSize,
        Packer,
        TextRun,
        Table,
        TableRow,
        TableCell,
        WidthType,
        AlignmentType,
        Numbering,

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
/*
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
*/
const makeTable = (rows) =>
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: rows.map((row, rowIndex) =>
        new TableRow({
          children: row.map((cell) =>
            new TableCell({
              width: { size: 20, type: WidthType.PERCENTAGE },
              shading: rowIndex === 0
                ? {
                    fill: "003366", // Dark blue background for header
                  }
                : undefined,
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cell,
                      font: "Montserrat",
                      size: 24,
                      color: rowIndex === 0 ? "FFFFFF" : undefined, // White text in header
                      bold: rowIndex === 0, // Make header text bold
                    }),
                  ],
                }),
              ],
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

    const makeBulletList = (items) => {
        return items.map((item) => {
            return new Paragraph({
            bullet: { level: 0 },
            children: [
                new TextRun({
                text: item,
                font: "Montserrat",
                size: 24, // 24 half-points = 12pt
                }),
            ],
            });
        });
    };
/*
    const makeItineraryTable = (rows) =>
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: rows.map((row) =>
            new TableRow({
              children: row.map((cell, colIndex) =>
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children:
                    colIndex === 2 && Array.isArray(cell) // If it's the itinerary column and it's an array
                      ? makeBulletList(cell) // Use bullet list
                      : [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: cell,
                                font: "Montserrat",
                                size: 24,
                              }),
                            ],
                          }),
                        ],
                })
              ),
            })
          ),
        });
*/

const makeItineraryTable = (rows) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map((row, rowIndex) =>
      new TableRow({
        children: row.map((cell, colIndex) => {
          // Set dynamic column widths: Itinerary column (index 2) gets more space
          let cellWidth;
          if (colIndex === 2) {
            cellWidth = 40; // Itinerary column
          } else {
            cellWidth = 15; // Other columns
          }

          const isHeader = rowIndex === 0;

          return new TableCell({
            width: { size: cellWidth, type: WidthType.PERCENTAGE },
            shading: isHeader
              ? {
                  fill: "003366", // Dark blue background for header
                }
              : undefined,
            children:
              colIndex === 2 && Array.isArray(cell)
                ? makeBulletList(cell)
                : [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [
                        new TextRun({
                          text: cell,
                          font: "Montserrat",
                          size: 24,
                          color: isHeader ? "FFFFFF" : undefined, // White text for dark blue header
                          bold: isHeader,
                        }),
                      ],
                    }),
                  ],
          });
        }),
      })
    ),
  });

    // Create a table for flight information
    const flightRows = [
        ["Flight Number", "Departure Date", "Departure Time", "Arrival Date", "Arrival Time"]
    ];
          
    // Add data rows from flightDetails
    flightDetails.forEach(flight => 
    {
        flightRows.push([
            flight.flightNum,
            flight.departureDate,
            flight.departureTime, 
            flight.arrivalDate,
            flight.arrivalTime    
        ]);
    });

    // Create a table for itinerary information
    const itineraryRows = [
        ["Day", "City", "Itinerary", "Meals", "Hotel Name"]
    ];

    // Add data rows from itineraryDetails
    itineraryDetails.forEach(itinerary => 
    {
        //Turn day to string
        const day = itinerary.day;
        //Convert to string
        itinerary.day = day.toString();

        //Check if meals are checked
        if (itinerary.bf) {
            itinerary.bf = "B";
        } else {
            itinerary.bf = "X";
        }
        if (itinerary.l) {
            itinerary.l = "L";
        } else {
            itinerary.l = "X";
        }
        if (itinerary.d) {
            itinerary.d = "D";
        } else {
            itinerary.d = "X";
        }

        itineraryRows.push([
            itinerary.day,
            itinerary.city,
            itinerary.descArray,
            itinerary.bf + "/" + itinerary.l + "/" + itinerary.d,
            itinerary.hotelName
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
            heading("TOUR VOUCHER"),
            new Paragraph(""),
            labelValue("Date of Issue", voucherDate),
            labelValue("Destination", destination),
            labelValue("Lead Name", customerName),
            labelValue("Date of Travel", travelDate),
            new Paragraph(""),
    
            heading("Hotel Info"),
            new Paragraph(""),
            labelValue("Hotel Name", hotelName),
            labelValue("Rooms", rooms),
            labelValue("Hotel Conf #", hotelConf),
            labelValue("Hotel Address", hotelAddress),
            new Paragraph(""),
            //Passenger Names
            heading("Passenger Names"),
            new Paragraph(""),
            ...makeBulletList(passengerNamesArray),
            new Paragraph(""),

            //Flight Information
            heading("FLIGHT INFORMATION"),
            new Paragraph(""),
            makeTable(flightRows),
            new Paragraph(""),
            
            //Itinerary Information
            heading("ITINERARY"),
            new Paragraph(""),
            makeItineraryTable(itineraryRows),
            new Paragraph(""),
            
            //SCOPE OF WORK
            heading("Inclusions"),
            new Paragraph(""),
            ...makeBulletList(incDescArray),
            new Paragraph(""),
    
            heading("Exclusions"),
            new Paragraph(""),
            ...makeBulletList(excDescArray),
            new Paragraph(""),

            heading("Arrival Instructions"),
            new Paragraph(""),
            ...makeBulletList(arrInsArray),
            new Paragraph(""),
    
            heading("Remarks"),
            new Paragraph(""),
            ...makeBulletList(remArray),
            new Paragraph(""),
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