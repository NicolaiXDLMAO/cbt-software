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

    const flightNumID = "flightNum" + inputFlightCount;
    const departureDateID = "departureDate" + inputFlightCount;
    const departureTimeID = "departureTime" + inputFlightCount;
    const arrivalDateID = "arrivalDate" + inputFlightCount;
    const arrivalTimeID = "arrivalTime" + inputFlightCount;

    const flightContainer = document.getElementById('flightContainer');
    const newFlightNum = document.createElement('input');
    const newFlightNumLabel = document.createElement('label');

    const newDepartureDate = document.createElement('input');
    const newDepartureDateLabel = document.createElement('label');
    const newDepartureTime = document.createElement('input');
    const newDepartureTimeLabel = document.createElement('label');

    const newArrivalDate = document.createElement('input');
    const newArrivalDateLabel = document.createElement('label');
    const newArrivalTime = document.createElement('input');
    const newArrivalTimeLabel = document.createElement('label');

    //Flight No.
    newFlightNumLabel.setAttribute('for', flightNumID);
    newFlightNumLabel.textContent = 'Flight No.: ';
    flightContainer.appendChild(newFlightNumLabel);

    newFlightNum.type = 'text';
    newFlightNum.placeholder = 'Flight No.';
    newFlightNum.id = flightNumID;
    newFlightNum.className = 'inputField';
    flightContainer.appendChild(newFlightNum);

    //Departure
    newDepartureDateLabel.setAttribute('for', departureDateID);
    newDepartureDateLabel.textContent = ' Departure Date: ';
    flightContainer.appendChild(newDepartureDateLabel);

    newDepartureDate.type = 'date';
    newDepartureDate.id = departureDateID;
    newDepartureDate.className = 'inputField';
    flightContainer.appendChild(newDepartureDate);

    newDepartureTimeLabel.setAttribute('for', departureTimeID);
    newDepartureTimeLabel.textContent = ' Departure Time: ';
    flightContainer.appendChild(newDepartureTimeLabel);

    newDepartureTime.type = 'time';
    newDepartureTime.id = departureTimeID;
    newDepartureTime.className = 'inputField';
    flightContainer.appendChild(newDepartureTime);

    //Arrival
    newArrivalDateLabel.setAttribute('for', arrivalDateID);
    newArrivalDateLabel.textContent = ' Arrival Date: ';
    flightContainer.appendChild(newArrivalDateLabel);

    newArrivalDate.type = 'date';
    newArrivalDate.id = arrivalDateID;
    newArrivalDate.className = 'inputField';
    flightContainer.appendChild(newArrivalDate);
    
    newArrivalTimeLabel.setAttribute('for', arrivalTimeID);
    newArrivalTimeLabel.textContent = ' Arrival Time: ';
    flightContainer.appendChild(newArrivalTimeLabel);

    newArrivalTime.type = 'time';
    newArrivalTime.id = arrivalTimeID;
    newArrivalTime.className = 'inputField';
    flightContainer.appendChild(newArrivalTime);
    flightContainer.appendChild(document.createElement("br"));

}

function moreItinerary()
{
    inputItineraryCount++;

    const dayID = "Day " + inputItineraryCount;
    const flightCodeID = "flightCode" + inputItineraryCount;
    const itiHotelID = "itiHotel" + inputItineraryCount;
    const bfID = "bf" + inputItineraryCount;
    const lID = "l" + inputItineraryCount;
    const dID = "d" + inputItineraryCount;
    const descID = "desc" + inputItineraryCount;

    const itineraryContainer = document.getElementById('itineraryContainer');
    const newDay = document.createElement('h3');
    const newFlightCode = document.createElement('input');
    const newFlightCodeLabel = document.createElement('label');
    const newItiHotel = document.createElement('input');
    const newItiHotelLabel = document.createElement('label');
    const newBf = document.createElement('input');
    const newBfLabel = document.createElement('label');
    const newL = document.createElement('input');
    const newLLabel = document.createElement('label');
    const newD = document.createElement('input');
    const newDLabel = document.createElement('label');
    const newDesc = document.createElement('textarea');
    const newDescLabel = document.createElement('label');

    //Days
    newDay.textContent = dayID;
    itineraryContainer.appendChild(newDay);

    //Flight Code
    newFlightCodeLabel.setAttribute('for', flightCodeID);
    newFlightCodeLabel.className = 'text';
    newFlightCodeLabel.textContent = 'Flight Code: ';
    itineraryContainer.appendChild(newFlightCodeLabel);

    newFlightCode.type = 'text';
    newFlightCode.placeholder = 'Flight Code';
    newFlightCode.id = flightCodeID;
    newFlightCode.className = 'inputField';
    itineraryContainer.appendChild(newFlightCode);

    //Hotel Name
    newItiHotelLabel.setAttribute('for', itiHotelID);
    newItiHotelLabel.className = 'text';
    newItiHotelLabel.textContent = ' Hotel Name: ';
    itineraryContainer.appendChild(newItiHotelLabel);

    newItiHotel.type = 'text';
    newItiHotel.placeholder = 'Hotel Name';
    newItiHotel.id = itiHotelID;
    newItiHotel.className = 'inputField';
    itineraryContainer.appendChild(newItiHotel);

    //Breakfast
    newBfLabel.setAttribute('for', bfID);
    newBfLabel.textContent = ' Breakfast';
    itineraryContainer.appendChild(newBfLabel);

    newBf.type = 'checkbox';
    newBf.id = bfID;
    newBf.className = 'checkboxField';
    itineraryContainer.appendChild(newBf);
    
    //Lunch
    newLLabel.setAttribute('for', lID);
    newLLabel.textContent = ' Lunch';
    itineraryContainer.appendChild(newLLabel);

    newL.type = 'checkbox';
    newL.id = lID;
    newL.className = 'checkboxField';
    itineraryContainer.appendChild(newL);

    //Dinner
    newDLabel.setAttribute('for', dID);
    newDLabel.textContent = ' Dinner';
    itineraryContainer.appendChild(newDLabel);

    newD.type = 'checkbox';
    newD.id = dID;
    newD.className = 'checkboxField';
    itineraryContainer.appendChild(newD);
    itineraryContainer.appendChild(document.createElement("br"));

    //Description
    newDescLabel.setAttribute('for', descID);
    newDescLabel.textContent = ' Description: ';
    itineraryContainer.appendChild(newDescLabel);

    newDesc.placeholder = 'Description';
    newDesc.id = descID;
    newDesc.className = 'inputField';
    newDesc.rows = 5;
    newDesc.cols = 50;
    itineraryContainer.appendChild(newDesc);
    itineraryContainer.appendChild(document.createElement("br")); // Add line break after each itinerary input
}

function generateVoucher() 
{
    //Get Header
    var headerImg = document.getElementById('header');

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
    const passengerNames = document.getElementById('passengerNames').value;

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
        const flightCode = document.getElementById('flightCode' + i).value;
        const hotelName = document.getElementById('itiHotel' + i).value;
        const bf = document.getElementById('bf' + i).checked ? 'Breakfast' : '';
        const l = document.getElementById('l' + i).checked ? 'Lunch' : '';
        const d = document.getElementById('d' + i).checked ? 'Dinner' : '';
        const desc = document.getElementById('desc' + i).value;
        itineraryDetails.push({day, flightCode, hotelName, bf, l, d, desc });
    }
    console.log(itineraryDetails);

    //Get Scope of Work
    const incDesc = document.getElementById('incDesc').value;
    const excDesc = document.getElementById('excDesc').value;
    const arrIns = document.getElementById('arrIns').value;
    const rem = document.getElementById('rem').value;

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
        ImageRun
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
            ["Day", "Flight Code", "Hotel Name", "Meals", "Description"]
        ];

    // Add data rows from itineraryDetails
    itineraryDetails.forEach(itinerary => 
    {
        //Turn day to string
        const day = itinerary.day;
        //Convert to string
        itinerary.day = day.toString();
        itineraryRows.push([
            itinerary.day,
            itinerary.flightCode,
            itinerary.hotelName,
            itinerary.bf + " " + itinerary.l + " " + itinerary.d,
            itinerary.desc    
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
                
            labelValue("Date of Issue", voucherDate),
            labelValue("Destination", destination),
            labelValue("Lead Name", customerName),
            labelValue("Date of Travel", travelDate),
    
            heading("Hotel Info"),
            labelValue("Hotel Name", hotelName),
            labelValue("Rooms", rooms),
            labelValue("Hotel Conf #", hotelConf),
            labelValue("Hotel Address", hotelAddress),
            labelValue("Names List", passengerNames),

            //Flight Information
            heading("FLIGHT INFORMATION"),
            makeTable(flightRows),
            
            //Itinerary Information
            heading("ITINERARY"),
            makeTable(itineraryRows),
            
            //SCOPE OF WORK
            heading("Inclusions"),
            makeParagraph(incDesc),
    
            heading("Exclusions"),
            makeParagraph(excDesc),

            heading("Arrival Instructions"),
            makeParagraph(arrIns),
    
            heading("Remarks"),
            makeParagraph(rem),
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
  