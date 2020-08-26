######################################
##  ------------------------------- ##
##     RG Template Generator        ##
##     Location Services Class      ##
##  ------------------------------- ##
##     Written by: David Shaeffer   ##
##  ------------------------------- ##
## Last Edited on: 12-20-2019       ##
##  ------------------------------- ##
######################################
"""This is the Location Services Class."""

from urllib.parse import urlencode
from urllib.request import urlopen
import json

class LocationServices:
    """This is the class doc string."""
    def __init__(self):
     """This takes the xml file path passed to the class and parses the xml using the lxml module.""" 
    #  self.xmlfile = ET.parse(xmlfile).getroot()

    def getWatershed(self, latitude, longitude):
        try:
            #This request takes a users lat/long and uses the raindrop point indexing service and return a comID
            params1 = urlencode([ ('pGeometry', 'POINT' + "(" + str(longitude) + " " + str(latitude) + ")"), ('pGeometryMod', 'WKT,SRSNAME=urn:ogc:def:crs:OGC::CRS84'), ('pPointIndexingMethod' , 'RAINDROP'), ( 'pPointIndexingMaxDist', '5'),
                                        ( 'pOutputPathFlag', 'FALSE'), ( 'pReturnFlowlineGeomFlag', 'FALSE'), ( 'optOutCS', 'SRSNAME=urn:ogc:def:crs:OGC::CRS84'), ( 'optOutPrettyPrint', '0'), ( 'optClientRef', 'CodePen'),
                                        ( 'f', 'json') ])    
            r1 = urlopen("https://ofmpub.epa.gov/waters10/PointIndexing.Service?" + params1)
            #print 'https://ofmpub.epa.gov/waters10/PointIndexing.Service?' + params1
            data = json.loads(r1.read())
            #pprint (data)
            comid = data['output']['ary_flowlines'][0]['comid']
            HUC8 = data['output']['ary_flowlines'][0]['wbd_huc12'] 
            global GNISName
            if data['output']['ary_flowlines'][0]['gnis_name'] is not None:
                GNISName = data['output']['ary_flowlines'][0]['gnis_name'] 
        except:
            raise
        try:
            #This request conducts a downstream main branch trace based on the comID from the point indexing service
            params2 = urlencode([  ('pNavigationType', 'DM'), ('pStartComid', comid), ('pStartMeasure', '100'), ('pTraversalSummary', 'TRUE'),('pFlowlinelist', 'TRUE'),('pEventList', ''), ('pEventListMod', ","), ('pStopDistancekm', "5"),
                                        ('optOutCS', 'SRSNAME=urn:ogc:def:crs:OGC::CRS84'), ('optOutPrettyPrint', '0'),('optClientRef', 'CodePen'),('f', 'json') ])
            r2 = urlopen("https://ofmpub.epa.gov/waters10/UpstreamDownstream.Service?" + params2)
            #print 'https://ofmpub.epa.gov/waters10/UpstreamDownstream.Service?' + params2
            flowdata = json.loads(r2.read())
            #pprint (flowdata)
            i=0
            try:
                # print (flowdata['output']['flowlines_traversed'][i])
                # print (flowdata)
                while flowdata['output']['flowlines_traversed'][i] is not None:
                    if flowdata['output']['flowlines_traversed'][i]['gnis_id'] is not None:
                        #This request gets the watershed characterics based on a comID
                        params3 = urlencode([ ('pComID', flowdata['output']['flowlines_traversed'][i]['comid'])])
                        #print 'https://ofmpub.epa.gov/waters10/nhdplus.jsonv25?' + params3
                        r3 = urlopen("https://ofmpub.epa.gov/waters10/nhdplus.jsonv25?" + params3)
                        watersheddata = json.loads(r3.read())
                        #print watersheddata
                        #global GNISName
                        GNISName = watersheddata['output']['header']['attributes'][0]['value']
                    i += 1
            except:
                # raise
                #do nothing 
                pass
        except:
            raise
            # pass
        return str(HUC8[0:8]) + "," + str(GNISName) 

    def getQuad(self, latitude, longitude):
              try:
                  #This request takes a users lat/long and uses the raindrop point indexing service and return a comID
                  params5 = urlencode([('geometry', '"x":' + str(longitude) + "," + '"y":' + str(latitude)), ('geometryType', 'esriGeometryPoint'), ('inSR', '4269'),
                                              ('spatialRel', 'esriSpatialRelWithin'), ('returnGeometry', 'false'), ('returnTrueCurves', 'false'), ('returnIdsOnly', 'false'), ('returnCountOnly', 'false'), ('returnZ:', 'false'), ('returnM:', 'false'), ('returnDistinctValues', 'false'),
                                              ('f', 'pjson')])
                  r5 = urlopen("https://index.nationalmap.gov/arcgis/rest/services/USTopoAvailability/MapServer/0/query?" + params5)
                  data = json.loads(r5.read())
                  quad = data['features'][0]['attributes']['CELL_NAME']
              except:
                  pass
              return str(quad)
    def getBasin(self, latitude, longitude):
              basin =''
              try:
                  #This request takes a users lat/long and uses the raindrop point indexing service and return a comID
                  params6 = urlencode([('geometry', '"x":' + str(longitude) + "," + '"y":' + str(latitude)), ('geometryType', 'esriGeometryPoint'), ('inSR', '4269'),
                                              ('spatialRel', 'esriSpatialRelWithin'), ('outFields', 'NAME'), ('returnGeometry', 'false'), ('returnTrueCurves', 'false'), ('returnIdsOnly', 'false'), ('returnCountOnly', 'false'), ('returnZ:', 'false'), ('returnM:', 'false'), ('returnDistinctValues', 'false'),
                                              ('f', 'pjson')])
                  r6 = urlopen("https://hydro.nationalmap.gov/arcgis/rest/services/wbd/MapServer/3/query?" + params6)
                #   print(r6.read())
                  data = json.loads(r6.read())
                  basin = data['features'][0]['attributes']['name']
                #   print(data)
              except Exception as e:
                  print(e)

              return str(basin)
    def getProjectAreaInfo(self, latitude, longitude):
        i=0
        try:
            params4 = urlencode([ ('latlng', str(latitude) + "," + str(longitude)), ('sensor', 'false'), ('key', 'ENTERKEYHERE') ])
            r4 = urlopen("https://maps.googleapis.com/maps/api/geocode/json?" + params4)
            locationdata = json.loads(r4.read())
            while locationdata['results'][0]['address_components'][i] is not None:
                    i += 1
                    if locationdata['results'][0]['address_components'][i]['types'][0] == 'locality':
                        city = locationdata['results'][0]['address_components'][i]['long_name']
                    if locationdata['results'][0]['address_components'][i]['types'][0] == 'administrative_area_level_1':
                        state = locationdata['results'][0]['address_components'][i]['short_name']
                    if locationdata['results'][0]['address_components'][i]['types'][0] == 'administrative_area_level_2':
                        county = locationdata['results'][0]['address_components'][i]['short_name']
        except:
                #do nothing 
                pass

        return str(city)+ "," + str(state) + "," + str(county.split(' County')[0])
    
    def getSpecies(self, State, County):
        try:
            #This request finds the county FIPS code
            SPparams = urlencode([('format', 'json'), ('columns', '/species@cn,sn,status,desc,listing_date'), ('filter', "/species/current_range_county@name = '" + County + "'"), ('filter', "/species@status = 'Endangered' or /species@status = 'Threatened' or /species@status = 'Proposed Endangered' or /species@status = 'Proposed Threatened'"), ('filter', "/species/range_state@abbrev = '" + State + "'")])
            # print 'https://ecos.fws.gov/ecp/pullreports/catalog/species/report/species/export?' + SPparams
            SPrequest = urlopen("https://ecos.fws.gov/ecp/pullreports/catalog/species/report/species/export?" + SPparams)
            SPdata = json.loads(SPrequest.read())
            i=0
            string = ''
            print(SPdata)
            while SPdata['data'][i] is not None:
                species = "Name: " + SPdata['data'][i][0] + " (" + SPdata['data'][i][1]['value'] + ")" + " Status: " + SPdata['data'][i][2]
                string += species
                string += " "
                # print string
                i+=1  
        except Exception as e:
                #if the user presses cancel then pass
                # print str(e)
                if str(e) != 'list index out of range':
                    print (str(e))
                pass
        return string


locations = LocationServices()
# print (locations.getWatershed(35.1920, -80.6570))
# print (locations.getQuad(35.1920, -80.6570))
# print (locations.getBasin(35.1920, -80.6570))
# print (locations.getProjectAreaInfo(35.1920, -80.6570))
# print (locations.getSpecies("SC", "Richland"))