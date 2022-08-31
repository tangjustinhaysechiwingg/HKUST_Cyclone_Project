/********************************** LOAD MODIS AQUA AND TERRA DATA ***********************************/
var collectionAqua = ee.ImageCollection("MODIS/006/MYD11A2")
  .filterDate('2020-06-01', '2020-06-20')
  .sort('system:time_start');                         //sort images by date captured


var collectionTerra = ee.ImageCollection("MODIS/006/MOD11A1")
  .filterDate('2020-06-01', '2020-06-11')
  .filter(ee.Filter.calendarRange(6, 8, 'month')) //restrict to images in summer months
  .sort('system:time_start');                         //sort images by date captured

//This function clips images to NYC
var clip = function(image) {
  return image.clip(geometry);
};
collectionTerra = collectionTerra.map(clip);  //map the clipping function over the image collection
collectionAqua = collectionAqua.map(clip);  //map the clipping function over the image collection

//print the number of images in the collection
var countTerra = collectionTerra.size();
var countAqua = collectionAqua.size();

print('Number of Images in the Image Collection:', countAqua);
print('Number of Images in the the Image Collection:', countTerra)


var changeBandName = function(image) {
  var LST1km = image.select('LST_Day_1km').multiply(.02).subtract(273.15);
  var LST = LST1km.rename('LST');
  return image.addBands(LST);
};

collectionAqua = collectionAqua.map(changeBandName);
collectionTerra = collectionTerra.map(changeBandName);

/*********************************** DISPLAY LST MAP ************************************/
//ADDS LST MAP LAYER

//color palette for LST map visualization
var LSTpalette = ['blue', 'cornflowerblue', 'aqua', 'greenyellow',
                  'yellow', 'gold', 'orange', 'red'];

//Visualization Parameters for Map
var TsVisParams = {
    min: 20, max: 50,    //use with C
  //min: 80, max: 100,   // use with F
  palette: LSTpalette};

//Adds first image layer to map, do not display
Map.addLayer(collectionTerra.first().select('LST'), TsVisParams, 'Terra LST', 0);
Map.addLayer(collectionAqua.first().select('LST'), TsVisParams, 'Aqua LST', 0);

//print date of image displayed to console
var date = collectionAqua.first().get('system:time_start');
print('Date of Aqua Map Displayed:', ee.Date(date));

var date = collectionTerra.first().get('system:time_start');
print('Date of Terra Map Displayed:', ee.Date(date));

//Image of mean LST calculated over the image collection
var mean_Aqua_LST = collectionAqua.mean().select('LST');
var mean_Terra_LST = collectionTerra.mean().select('LST');

//Apply correct visualization parameters to mean LST Image
var mean_Aqua_image = mean_Aqua_LST.visualize(TsVisParams); // create the mean image
var mean_Terra_image = mean_Terra_LST.visualize(TsVisParams); // create the mean image

//Displays the average LST for each pixel over all the images on Map
Map.addLayer(mean_Aqua_image, {}, 'Mean Aqua LST');
Map.addLayer(mean_Terra_image, {}, 'Mean Terra LST');

/*********************************** CONDUCT CENSUS TRACT ANALYSIS ************************************/

// Establish each community district
var censusblocks = ee.FeatureCollection(table);

// For each community district, determine a mean LST value and graph.

print(ui.Chart.image.seriesByRegion(
  collectionTerra, censusblocks, ee.Reducer.mean(), 'LST', 30,'system:time_start', 'bctcb2010')
  );

print(ui.Chart.image.seriesByRegion(
  collectionAqua, censusblocks, ee.Reducer.mean(), 'LST', 30,'system:time_start', 'bctcb2010')
  );


Map.addLayer(censusblocks, {}, 'Census Tracts');
