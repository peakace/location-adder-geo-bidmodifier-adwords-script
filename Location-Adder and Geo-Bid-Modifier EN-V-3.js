/*
 *
 * Location-Adder and Geo-Bid-Modifier
 *
 * This script searches for target locations (based on your filter settings) in your campaigns, adds those locations 
 * and sets a bid modifier for them based on their conversion rates.
 * It also sets a bidding modifier for already existing locations. 
 * All changes made by the script will be reported in a report spreadsheet.
 * Created By: Marcel Prothmann, Alexander Tissen, Daniel Bartonitschek
 * Version: 1.2
 * www.pa.ag
 *
 */
 
//Filter Settings:
//This is the data time frame that the script will use to calculate the modifiers. Possible variations which can be used:
//TODAY, YESTERDAY, LAST_7_DAYS, THIS_WEEK_SUN_TODAY, LAST_WEEK, LAST_14_DAYS, LAST_30_DAYS, LAST_BUSINESS_WEEK, LAST_WEEK_SUN_SAT, THIS_MONTH, LAST_MONTH, ALL_TIME
var DATE = "LAST_30_DAYS";

// This is the URL of the spreadsheet containing all possible geo target locations (and IDs) worldwide. 
// You can use our spreadsheet or just add your own location spreadsheet (e.g. https://docs.google.com/spreadsheets/Ihre_Location_URL). 
// All the supported location can be downloaded here: https://developers.google.com/adwords/api/docs/appendix/geotargeting
var LOCATIONS_URL = 'https://docs.google.com/spreadsheets/d/1-boBTeruZBAecNItuom5aMRqFqpYQ3zSUm3a5xKCg3E/';

//This is the name of the sheet with the spreadsheet containing the locations.
var LOCATIONS_SHEET_NAME = 'locations';

//This is the URL of the spreadsheet where you will see the daily changes of the script.
// Please just create a new empty spreadsheet and replace the URL here:
var REPORTING_URL = 'https://docs.google.com/spreadsheets/Ihre_Reporting_URl';

//This is the minimum amount of clicks (within the time frame) the locations must have in order to be considered for bidding
var MIN_CLICKS = 1;

//This is the minimum amount of impressions (within the time frame) the locations must have in order to be considered for bidding
var MIN_IMPRESSIONS = 1;

//This is the minimum amount of costs (within the time frame) the locations must have in order to be considered for bidding
var MIN_COST = 0;

//These are the minimum and maximum possible bid-modifiers. You can set these between 0.1 and 10 ( 0.1 = -90% , 10 = +900% )
var MAX_BID = 3
var MIN_BID = 0.5

//This is for campaigns which should not be checked. Please separate them with a comma.
var EXCLUDE_CAMPAIGNS = ["Campaign Example 1", "Campaign Example 2"];

//This is the minimum of amount of clicks the new locations must have in order to be considered for bidding
var MIN_LOCATION_CLICKS = 50;


function findLastCell( sheet, column ){
  var lastCell = 1;
  
  while( sheet.getRange( lastCell, column ).getValue() != "" ){
    lastCell = lastCell + 1;
  }
  return lastCell;
}

/*
 * Get the day of week: monday = 1, tuesday = 2, ...
*/
function dayOfWeek(){
  var localOffsetInHours = 2;
  var date = new Date();
  var offsetInMinutes = date.getTimezoneOffset() + localOffsetInHours * 60;
  var d = new Date(date.valueOf() + offsetInMinutes * 60 * 1000);
  
  var day = d.getDay();
  
  if( day == 0 ){
    // we need it monday-based, not sunday-based
    day = 7;
  }
  return day;
}


/*
 * Each campaign has its own sheet in the workbook. 
 * return the sheet for the given campaignName, create it if is missing
*/
function getSheet( workBook, campaignName ){
  var sheet = workBook.getSheetByName(campaignName);
  if( ! sheet ){
    sheet = workBook.insertSheet(campaignName);
    sheet.getRange('A1').setValue("Montag");
    sheet.getRange('B1').setValue("Dienstag");
    sheet.getRange('C1').setValue("Mittwoch");
    sheet.getRange('D1').setValue("Donnerstag");
    sheet.getRange('E1').setValue("Freitag");
    sheet.getRange('F1').setValue("Samstag");
    sheet.getRange('G1').setValue("Sonntag");
  }
  
  return sheet;
}

function main(){
  var reportingBook = SpreadsheetApp.openByUrl(REPORTING_URL);
  
  var sheets = reportingBook.getSheets();
  
  // header row
  var weekDay = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag","Sonntag"];
  
  // in all sheets (for each campaign) clear the column which corresponds to the current weekday
  for( index in sheets ){
    var sheet = sheets[index];
    var column = dayOfWeek();
    sheet.deleteColumn( column );
    sheet.insertColumnBefore( column );
    sheet.getRange( 1, column ).setValue( weekDay[column-1] );
    sheet.autoResizeColumn(column);
  }
  
  // read all locations from sheet and store them in a map: canonical-name -> location-id
  var locations = getLocations();
 
  
  // add new targeted locations to all campaigns
 var campaignIterator = AdWordsApp.campaigns()
	.withCondition("Status = 'ENABLED'")
	.get(); 
 while( campaignIterator.hasNext() ){
    var campaign = campaignIterator.next();
    
    if( EXCLUDE_CAMPAIGNS.indexOf(campaign.getName()) >= 0 ){
      continue;
    }
    
    addNewLocationsWithBidModifier( campaign, locations, reportingBook );
  }
  
  // update bids for existing targeted locations
  Logger.log("und jetzt die getargeteten Locations");
  updateBidModifierForexistingLocations( reportingBook );
}



function addNewLocationsWithBidModifier( campaign, locations, reportingBook ){
  
  // get metrics for all locatioins with conversions > 1
  var report = AdWordsApp.report(
    "SELECT CampaignId, RegionCriteriaId, CountryCriteriaId, CityCriteriaId, Clicks, Conversions " +
    "FROM   GEO_PERFORMANCE_REPORT " +
    "WHERE  Clicks >= " + MIN_CLICKS + " " +
    "AND CampaignId IN [" + campaign.getId() + "] " + 
    "AND Conversions > 1 " +
      "AND Impressions >= " + MIN_IMPRESSIONS + " " +
        "DURING " + DATE );
  
  var rows = report.rows();
  
  clicksMap = {};
  conversionsMap = {};
  cityMap = {};  
  while (rows.hasNext()) {
    var row = rows.next();
    var location = row["CountryCriteriaId"];
    var city = row["CityCriteriaId"];
    var region = row["RegionCriteriaId"];    
    var id = locations[ city + "," + region + "," + location ];
    
    if( ! id ){
      // if location-id is not found, then skip this step
      continue;
    }
    
    cityMap[ id ] = city + "," + region + "," + location;
  
    clicksMap[id] = clicksMap[id] ? clicksMap[id] : 0;
    conversionsMap[ id ] = conversionsMap[ id ] ? conversionsMap[ id ] : 0;
    
    clicksMap       [ id ] = clicksMap[ id ] + row["Clicks"];
    conversionsMap[ id ] = conversionsMap[ id ] + row["Conversions"];
    
  }
  
  
  
  
  loop :  for( var id in clicksMap ){
    var clicks = clicksMap[id];
    var conversions = conversionsMap[id];
    var conversionrate = conversions / clicks * 100;
       
    var campaincvr = campaign.getStatsFor( DATE ).getConversionRate(); 
    var bid_modifier = (conversionrate / campaincvr) / 100;
    
    
    
    
    bid_modifier = Math.max( bid_modifier, MIN_BID );
    bid_modifier = Math.min( bid_modifier, MAX_BID );
    
    if( bid_modifier != bid_modifier ){
      // if bid_modifier is NaN then ignore this location
      continue;
    }
    
    var locationIterator = campaign.targeting().targetedLocations().get();
    while( locationIterator.hasNext() ){
      var targetedLocation = locationIterator.next();
      if( targetedLocation.getId() == id ){
		  // ignore locations which are already targeted here. they will be processed later
        continue loop;
      }
    }
    
    
    campaign.addLocation( Math.round( id ), bid_modifier );
    
    
    var reportingSheet = getSheet( reportingBook, campaign.getName() );  
    var city = cityMap[id];
    var column = dayOfWeek();
    Logger.log("column: " + column );
    var row = findLastCell( reportingSheet, column );
    var log = city + " (" + id + ") : " + ( Math.round( ( bid_modifier - 1) * 100 ) + "%" );
    reportingSheet.getRange(row, column).setValue( log );
    reportingSheet.autoResizeColumn(column);
    
    Logger.log (city + "," + region + "," + location + ' (' + id + ') ' + "Klicks: " + clicks + " Conversions: " + conversions);
  }
  
}



function updateBidModifierForexistingLocations( reportingBook ){
  
  var campaignIterator = AdWordsApp.campaigns()
  .withCondition("Status = 'ENABLED'")
  .withCondition("Conversions > 0")
  .withCondition("Clicks > " + MIN_CLICKS)
  .forDateRange(DATE)
  .get();
  
	while( campaignIterator.hasNext() ){
		var campaign = campaignIterator.next();	
		if( EXCLUDE_CAMPAIGNS.indexOf(campaign.getName()) >= 0 ){
		  continue;
		}
		
		var campaincvr = campaign.getStatsFor( DATE ).getConversionRate();
		//Logger.log(campaincvr);
		var locationIterator = campaign.targeting().targetedLocations().get();
		while( locationIterator.hasNext() ){
			var targetedLocation = locationIterator.next(); 
			var stats = targetedLocation.getStatsFor( DATE );
			var locationClicks = stats.getClicks();
				if( locationClicks < MIN_LOCATION_CLICKS ){
				continue;
			}
				
			var locationcvr = stats.getConversionRate();
			var bid = locationcvr / campaincvr;
			 
			if (bid > MAX_BID) {
				bid = MAX_BID;
			} 
			if (bid < MIN_BID) {
				bid = MIN_BID;
			}

			if( bid != bid ){
				// if bid_modifier is NaN then ignore this location
				continue;
			} 
			
			var oldBid = targetedLocation.getBidModifier();
			//  Logger.log(oldBid)
			if( Math.abs( oldBid - bid ) < .002 ){
				// if bid-modifier roughly equals the old bid-modifier then ignore this
				continue;     
			}
			 
			targetedLocation.setBidModifier( bid );
			
			var reportingSheet = getSheet( reportingBook, campaign.getName() );
			var column = dayOfWeek();
			//Logger.log("column: " + column );
			var row = findLastCell( reportingSheet, column );
			var log = targetedLocation.getName() + " : " + ( Math.round( ( bid - 1 ) * 100 ) + "%" );
			reportingSheet.getRange( row, column ).setValue( log );
			reportingSheet.autoResizeColumn(column);
		}
	}
}


function getLocations(){
  var ss = SpreadsheetApp.openByUrl( LOCATIONS_URL );
  var sheet = ss.getSheetByName( LOCATIONS_SHEET_NAME );
  
  var map = [];
  
  // This represents ALL the data.
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  // This logs the spreadsheet in CSV format.
  for( var i = 0; i < values.length; i++ ){
    // map[ city_name ] = id;
    // 1 = zweite Spalte (city_name)
    // 0 = erste Spalte ( id )
    map[ values[i][2] ] = values[i][0];
  }
  
  return map;
}
