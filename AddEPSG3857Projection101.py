#ArcGIS Server 10.1 service editor
#view your service properties at: http://[your server URL]/arcgis/admin/services/
#put a ?f=json at the end of a service name to see the json properties -
#the JSON is what is being edited here
#Loops through the services in a particular folder and edits the
#listSupportedCRS property (adds "EPSG:3857" - to be google-riffic) for each WMS service in the folder
#created by Doug Curl, Kentucky Geological Survey, 9/12/2013

# For HTTP calls
import httplib, urllib, json

# For system tools
import sys

# For reading passwords without echoing
import getpass

def main(argv=None):

    # Ask for server name & port
    #serverName = "kgs.uky.edu"
    serverName = raw_input("Enter server name (server URL): ")
    # Ask for server port - usually 6080:
    serverPort = raw_input("Enter server port (usually 6080): ")
    #Ask for server admin directory:
    serverFolder = raw_input("Enter folder in your service directory to edit services (assumes the root is '/arcgis/admin/services/'): ")

    # Ask for admin/publisher user name and password
    username = raw_input("Enter admin user name: ")
    password = getpass.getpass("Enter password: ")

    # Get a token
    token = getToken(username, password, serverName, serverPort)

    # Get the root info
    #serverURL = "/arcgis/admin/services/aasggeothermal/"
    serverURL = "/arcgis/admin/services/"+serverFolder+"/"

    # This request only needs the token and the response formatting parameter 
    params = urllib.urlencode({'token': token, 'f': 'json'})

    headers = {"Content-type": "application/x-www-form-urlencoded", "Accept": "text/plain"}

    # Connect to URL and post parameters    
    httpConn = httplib.HTTPConnection(serverName, serverPort)
    httpConn.request("POST", serverURL, params, headers)

    # Read response
    response = httpConn.getresponse()
    if (response.status != 200):
        httpConn.close()
        print "Could not read folder information."
        return
    else:
        data = response.read()
        
        # Check that data returned is not an error object
        if not assertJsonSuccess(data):          
            print "Error when reading server information. " + str(data)
            return
        else:
            print "Processed server information successfully. Now processing folders..."

        # Deserialize response into Python object
        dataObj = json.loads(data)
        httpConn.close()

        # Loop through each service in the folder   
        for item in dataObj['services']:
            print item["serviceName"]
            print item["type"]
            if item["type"] == "MapServer":
                service = item["serviceName"]+"."+item["type"]
                #sUrl = "/arcgis/admin/services/%s.%s" %(item["serviceName"], item["type"])
                print service
                serviceURL = serverURL + service
                print serviceURL
                # This request only needs the token and the response formatting parameter 
                params = urllib.urlencode({'token': token, 'f': 'json'})
                
                headers = {"Content-type": "application/x-www-form-urlencoded", "Accept": "text/plain"}
                
                # Connect to service to get its current JSON definition    
                httpConn = httplib.HTTPConnection(serverName, serverPort)
                httpConn.request("POST", serviceURL, params, headers)
                
                # Read response
                response = httpConn.getresponse()
                if (response.status != 200):
                    httpConn.close()
                    print "Could not read service information."
                    return
                else:
                    data = response.read()

                    # Check that data returned is not an error object
                    if not assertJsonSuccess(data):          
                        print "Error when reading service information. " + str(data)
                    else:
                        print "Service information read successfully. Now changing properties..."

                    # Deserialize response into Python object
                    dataObj = json.loads(data)

                    httpConn.close()
                    #print data

                    for ext in dataObj["extensions"]:
                        if ext["typeName"] == "WMSServer":
                            #Edit the supported CRS property - add the one for google for WMS:
                            ext["properties"]["listSupportedCRS"] = "EPSG:3857"
                    
                    # Serialize back into JSON
                    updatedSvcJson = json.dumps(dataObj)
                    #print updatedSvcJson
                    # Call the edit operation on the service. Pass in modified JSON.
                    editSvcURL = serverURL + service + "/edit"
                    params = urllib.urlencode({'token': token, 'f': 'json', 'service': updatedSvcJson})
                    httpConn.request("POST", editSvcURL, params, headers)
                    
                    # Read service edit response
                    editResponse = httpConn.getresponse()
                    if (editResponse.status != 200):
                        httpConn.close()
                        print "Error while executing edit."
                        return
                    else:
                        editData = editResponse.read()
                        
                        # Check that data returned is not an error object
                        if not assertJsonSuccess(editData):
                            print "Error returned while editing service" + str(editData)        
                        else:
                            print "Service edited successfully."

                   #httpConn.close()  
                    #return
            else:
              # Close the connection to the current service
              httpConn.close()

# A function to generate a token given username, password and the adminURL.

def getToken(username, password, serverName, serverPort):
    # Token URL is typically http://server[:port]/arcgis/admin/generateToken
    tokenURL = "/arcgis/admin/generateToken"
    
    params = urllib.urlencode({'username': username, 'password': password, 'client': 'requestip', 'f': 'json'})
    
    headers = {"Content-type": "application/x-www-form-urlencoded", "Accept": "text/plain"}
    
    # Connect to URL and post parameters
    httpConn = httplib.HTTPConnection(serverName, serverPort)
    httpConn.request("POST", tokenURL, params, headers)
    
    # Read response
    response = httpConn.getresponse()
    if (response.status != 200):
        httpConn.close()
        print "Error while fetching tokens from admin URL. Please check the URL and try again."
        return
    else:
        data = response.read()
        httpConn.close()
        
        # Check that data returned is not an error object
        if not assertJsonSuccess(data):            
            return
        
        # Extract the token from it
        token = json.loads(data)        
        return token['token']            
        

# A function that checks that the input JSON object 
#  is not an error object.
    
def assertJsonSuccess(data):
    obj = json.loads(data)
    if 'status' in obj and obj['status'] == "error":
        print "Error: JSON object returns an error. " + str(obj)
        return False
    else:
        return True
    
        
# Script start 
if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
