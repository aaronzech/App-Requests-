// EMAIL_SENT is written in column P for rows for which an emails have been sent successfully
// This is done so that repeated emails don't get sent

//Global Variables
var EMAIL_SENT = "EMAIL_SENT"; // checks the cell value
var paidApp= "Paid";


function sendEmails2()
{
  
  var sheet = SpreadsheetApp.getActiveSheet(); // Grabs the spreadsheet
  var startRow = sheet.getLastRow();;  // Starts at the last row of data to process
  var numRows = 1;   // Number of rows to process, (Only doing one row, because we are working on the last row)
  
  // Creates an array of what cells to grab
  var dataRange = sheet.getRange(startRow, 1, numRows, 18) // grabs collumns A-P 
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) 
  {
      var row = data[i]; 
      var emailAddress = row[1];  // Points to Email column
      var message = "00000000000000000";  
      var emailSent = row[20]; // Points to the last last in column R
      var appPaid = row[7]; // Row H
    
    
      if (appPaid != paidApp) // Send this message if it is a free app
      {
         message ="*This is an automatic reply*" + "\n" + "\n" 
      + "Thank you for submitting your request for " + row[0] 
      + ". The app you recommended will be reviewed in detail, and be presented at the next Student Device App Committee meeting, which takes place monthly." 
      +'\n'+'\n' 
      + "You will receive another follow up email, explaining whether or not the app you have recommended was approved or denied. If additional information is needed to make a decision on the app that you have requested, you will also be contacted.";       
  
      }
      else  // Send this message if it is a paid app
      {
        message ="*This is an automatic reply*" + "\n" + "\n"
        + "Thank you for submitting your request for " + row[0] 
        + ". The app you have recommended  is a paid app, at this time Osseo Area Schools is not considering paid applications.";
        
        // SET ROW COLOR TO LIGHT RED
        for (var a = 1; a<17; ++a)
        {
          sheet.getRange(startRow, a).setBackgroundColor("#ff6666");
        }
      
      }
    
   
     
    
    if (emailSent != EMAIL_SENT) // Prevents sending duplicates
    {  
          var subject = "Self-Service App Recommendation";
      
         
      
          MailApp.sendEmail( emailAddress, subject, message); // email gets sent to this address
      
          sheet.getRange(startRow, 17).setValue(EMAIL_SENT); //+i
          // Make sure the cell is updated right away in case the script is interrupted
          populateSlide();
          SpreadsheetApp.flush();
    }
  }
  
}

// Create Slide Deck
// This populates a slide deck with the information submitted via a Google form
function populateSlide() {

  var sheet = SpreadsheetApp.getActiveSheet(); // Grabs the spreadsheet
  var startRow = sheet.getLastRow();;  // Starts at the last row of data to process
  var numRows = 1;   // Number of rows to process, (Only doing one row, because we are working on the last row)
  // Creates an array of what cells to grab
  var dataRange = sheet.getRange(startRow, 1, numRows, 16) // grabs collumns A-P 
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) 
  {
      var row = data[i]; 
      var jobTitle = row[2];  // Points to Email column
      var name = row[3];
      var description = row[5];
      var appLink = row[6];
      var why = row[9];
      var curric = row[10];
      var appName = row[0];
      var platform = row[12]; // What Platform is the App On?
  }
 

  
  
  // Create 4 new slides at the end of Slide Deck
  var slideDeck =SlidesApp.openByUrl("https://docs.google.com/presentation/d/1FUp5WiK2-RL8ygRbU5zCiUBpYnYMkFaAQ9fqCOhQut0/edit#slide=id.g30d771fc9c_0_82");
 
  // COPY TEMPLATE FROM SLIDE#3
  var slideOne = slideDeck.getSlides()[2];
  var slideOne = slideDeck.appendSlide(slideOne);
  
  var lastSlideIndex = slideDeck.getSlides().length;
  Logger.log(lastSlideIndex);
  for(var i = lastSlideIndex; i<lastSlideIndex+2; i++)
  {
    slideDeck.insertSlide(i);
  }
 
  
  // COPY TEMPLATE FROM SLIDE#2
  var slideFour = slideDeck.getSlides()[1];
  var slideFour = slideDeck.appendSlide(slideFour);

 
  // SLIDE 1 
  //-----------------------------------------------------------------------------------------------------------------------
  
  //TITLE SLIDE - NEW
  // TITLE TEXT                                                       Location X, Location Y, Width,height
  var titleshape = slideOne.insertShape(SlidesApp.ShapeType.TEXT_BOX, 10, 60, 450, 200);
  var textRange = titleshape.getText();
  textRange.setText(appName);
  textRange.getTextStyle().setBold(true).setFontSize(36).setFontFamily("Merriweather");
  
  // REQUESTOR NAME & TITLE                                      Location X, Location Y, Width,height
  var shape = slideOne.insertShape(SlidesApp.ShapeType.TEXT_BOX, 10, 250, 300, 60);
  var textRange = shape.getText();
  textRange.setText(
    "Name: " + name + "\n"+
    "Title: " + jobTitle);
  textRange.getTextStyle().setFontSize(16) // Font size
  shape.getFill().setSolidFill("#ffffff", .8);
 // shape.getFill().setSolidFill(color, alpha)
 
 
 
  //Platform                                                  Location X, Location Y, Width,height
  var platformBox = slideOne.insertShape(SlidesApp.ShapeType.TEXT_BOX, 600, 18, 130, 30);
 
  if(platform=="iOS"){
     
    createiOSLogo(slideOne);
    
  }
  if(platform=="Android"){
   
    createAndroidLogo(slideOne);
  }
  if(platform=="Chrome Web Store Extension"){
  
    createChromeLogo(slideOne);
    formatShape(platformBox,200,60,500,10);    
  }    
  // ios & Android
  if(platform=="iOS, Android"){
       createiosAndroidLogos(slideOne);
  }
  
 
  var textRange = platformBox.getText();
  textRange.setText("Platform: "+platform);
  textRange.getTextStyle().setForegroundColor("#ffffff");
  
 
  
  // SLIDE 2
  // INCERT APP DESCRIPTION AND LINK
  //----------------------------------------------------------------------------------------------------------------
  var slide2 = slideDeck.getSlides()[lastSlideIndex];
  
  var shape = slide2.insertShape(SlidesApp.ShapeType.TEXT_BOX, 10, 130, 400, 200);
 
  var textRange = shape.getText();
  textRange.setText(" " + description + "\n");
  textRange.getTextStyle().setFontSize(16); // Change font to size 16
  //var descriptionText = textRange.appendText("Description:");
  //descriptionText.getTextStyle().setBold(true);
  var descriptionText = textRange.insertText(0,"Description:")
  descriptionText.getTextStyle().setBold(true).setFontFamily("Merriweather").setFontSize(18);
  

  var insertedText = textRange.appendText("App Link");
  insertedText.getTextStyle()
    .setBold(true)
    .setLinkUrl(appLink)
    .setForegroundColor('#ff0000');
  
  // INSERT IMAGE
https://images.techhive.com/images/article/2013/05/tablet_cat-100038149-large.jpg
  //----------------------------------------------------------------------------------------------------------------------
  
  
  //WHY DO YOU RECCOMEND THIS APP 
  //SLIDE #3
  //----------------------------------------------------------------------------------------------------------------------
  var whySlide = slideDeck.getSlides()[lastSlideIndex+1]; // Assigning to slide 3
  var shape = whySlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 15, 15, 300, 40);
  //   Add Text box                                         left, top, box width, box heigh
  var textRange = shape.getText();
  // SET TEXT
  textRange.setText("Why do you recommend this app?\n");
  // FORMAT TEXT
  textRange.getTextStyle().setBold(true).setFontFamily("Merriweather").setFontSize(16);
  var shape1 = whySlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 15, 55, 300, 80);
  //       Add Text box                                               left, top, box width, box heigh
  var textRange = shape1.getText();
  textRange.setText(why);
  
 
  var shape2 = whySlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 15, 200, 300, 40);
  var textRange = shape2.getText();
  // SET TEXT
  textRange.setText("How would you use this app in your curriculum? \n");
  // FORMAT TEXT
  textRange.getTextStyle().setBold(true).setFontFamily("Merriweather").setFontSize(16);
  var shape3 = whySlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 15, 240, 300, 80);
  //                                                      left, top, box width, box heigh
  var textRange = shape3.getText();
  textRange.setText(curric);
  
  //------------------------------------------------------------------------------------------------------------------------
  
}

function createChromeLogo(slide)
{
  imageURL = "https://storage.googleapis.com/gweb-uniblog-publish-prod/images/Chrome__logo.max-500x500.png"
  Chromelogo = slide.insertImage(imageURL);
  Chromelogo.setWidth(50);
  Chromelogo.setHeight(50);
  Chromelogo.setTop(350);
  Chromelogo.setLeft(10);
}
function createiOSLogo(slide)
{
  imageURL = "https://cdn0.iconfinder.com/data/icons/flat-round-system/512/iOS-512.png"
  iosLogo = slide.insertImage(imageURL);
  iosLogo.setWidth(50);
  iosLogo.setHeight(50);
  iosLogo.setTop(350);
  iosLogo.setLeft(10);
}
function createAndroidLogo(slide)
{
  imageURL = "https://cdn1.iconfinder.com/data/icons/logotypes/32/android-512.png"
  iosLogo = slide.insertImage(imageURL);
  iosLogo.setWidth(50);
  iosLogo.setHeight(50);
  iosLogo.setTop(350);
  iosLogo.setLeft(10);
}
// Add Chrome image to the slides bottome left corner
function createiosChromeLogos(slide)
{
  imageURL = "https://cdn0.iconfinder.com/data/icons/flat-round-system/512/iOS-512.png"
  iosLogo = slide.insertImage(imageURL);
  iosLogo.setWidth(50);
  iosLogo.setHeight(50);
  iosLogo.setTop(350);
  iosLogo.setLeft(10);
  imageURL = "https://storage.googleapis.com/gweb-uniblog-publish-prod/images/Chrome__logo.max-500x500.png"
  Chromelogo = slide.insertImage(imageURL);
  Chromelogo.setWidth(50);
  Chromelogo.setHeight(50);
  Chromelogo.setTop(350);
  Chromelogo.setLeft(70);
}

// Add IOS image and Android Image to the slides bottom left corner
function createiosAndroidLogos(slide)
{
  imageURL = "https://cdn0.iconfinder.com/data/icons/flat-round-system/512/iOS-512.png"
  iosLogo = slide.insertImage(imageURL);
  iosLogo.setWidth(50);
  iosLogo.setHeight(50);
  iosLogo.setTop(350);
  iosLogo.setLeft(10);
  imageURL = "https://cdn1.iconfinder.com/data/icons/logotypes/32/android-512.png"
  Chromelogo = slide.insertImage(imageURL);
  Chromelogo.setWidth(50);
  Chromelogo.setHeight(50);
  Chromelogo.setTop(350);
  Chromelogo.setLeft(70);
}

// Adjust the size and location of the shape object
function formatShape(shape, width, height, left,top)
{
    shape.setWidth(width);
    shape.setHeight(height);
    shape.setLeft(left);
    shape.setTop(top);
    return shape;
}