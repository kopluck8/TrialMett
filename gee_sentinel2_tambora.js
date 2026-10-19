// Komposit Sentinel-2 musim kemarau + DEM untuk Savana Tambora (Google Earth Engine Code Editor).
// Tempel di https://code.earthengine.google.com lalu Run. Hasil diekspor ke Google Drive:
//   s2_kemarau_tambora.tif  (6 band: B2,B3,B4,B8,B11,B12, reflektan x10000, EPSG:32750, 10 m)
//   dem_tambora.tif         (Copernicus DEM GLO-30, di-resample ke grid yang sama)
// Di aplikasi (tab 5): Sentinel-2 = s2_kemarau_tambora.tif, urutan band = B2,B3,B4,B8,B11,B12,
// raster lain = dem_tambora.tif (+ slope/aspect dari ArcGIS Pro jika mau).

// ---- Atur di sini -------------------------------------------------------------
// Batas kira-kira Taman Nasional Tambora. Ganti dengan poligon kawasan/area studi Anda
// (Assets > Upload shapefile, lalu: var area = ee.FeatureCollection('users/.../batas').geometry();)
var area = ee.Geometry.Rectangle([117.75, -8.45, 118.25, -8.05]);
var tahun = 2026;
var awal = ee.Date.fromYMD(tahun, 7, 1);    // puncak kemarau, sesuaikan dengan waktu survei plot
var akhir = ee.Date.fromYMD(tahun, 10, 31);
var maksAwan = 40;                          // % awan per scene
// -------------------------------------------------------------------------------

function maskAwan(img) {
  // SCL: 3 bayangan awan, 8-10 awan & cirrus, 11 salju
  var scl = img.select('SCL');
  var ok = scl.neq(3).and(scl.neq(8)).and(scl.neq(9)).and(scl.neq(10)).and(scl.neq(11));
  return img.updateMask(ok);
}

var s2 = ee.ImageCollection('COPERNICUS/S2_SR_HARMONIZED')
  .filterBounds(area)
  .filterDate(awal, akhir)
  .filter(ee.Filter.lt('CLOUDY_PIXEL_PERCENTAGE', maksAwan))
  .map(maskAwan)
  .select(['B2', 'B3', 'B4', 'B8', 'B11', 'B12']);

print('Jumlah scene:', s2.size());
var komposit = s2.median().clip(area).toUint16();

var dem = ee.ImageCollection('COPERNICUS/DEM/GLO30').select('DEM').mosaic()
  .clip(area).toFloat();

Map.centerObject(area, 11);
Map.addLayer(komposit, {bands: ['B4', 'B3', 'B2'], min: 200, max: 2500}, 'RGB kemarau');
Map.addLayer(komposit.normalizedDifference(['B8', 'B4']), {min: 0, max: 0.7,
  palette: ['brown', 'yellow', 'green']}, 'NDVI');

Export.image.toDrive({
  image: komposit, description: 's2_kemarau_tambora', fileNamePrefix: 's2_kemarau_tambora',
  region: area, scale: 10, crs: 'EPSG:32750', maxPixels: 1e10
});
Export.image.toDrive({
  image: dem, description: 'dem_tambora', fileNamePrefix: 'dem_tambora',
  region: area, scale: 10, crs: 'EPSG:32750', maxPixels: 1e10
});
// Catatan: DEMNAS (BIG, 8 m) lebih detail daripada GLO-30; jika memakainya, cukup
// masukkan file DEMNAS ke aplikasi, aplikasi akan me-resample ke grid Sentinel-2.
