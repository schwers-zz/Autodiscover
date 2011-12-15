import httplib, urllib2
import base64
import sys
from ntlm import HTTPNtlmAuthHandler
import xml # for isinstance calls
from xml.dom.minidom import parseString

# Useful for Testing
START = "2010-12-12T00:00:00-08:00"
END = "2011-12-19T00:00:00-08:00"

# Set for Debug purposes
httplib.HTTPConnection.debuglevel = 1

# Define Helper Functions

def autodiscover_xml(email):
    """ Expects to have a string thats a valid email address for the intended exchange user
        Returns an xml string used to query the autodiscover servcie.
    """ 

    return """<?xml version="1.0" encoding="utf-8"?><Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006"><Request><EMailAddress>""" + email + """</EMailAddress><AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema></Request></Autodiscover>"""


def calendaritem_xml(shape, start, end):
    """ start & end -- valid utc strings designating a time range in which to query for 
            calendar events. 
        shpae -- xml string for the calendaritem shape according to the finditem spec, 
            this is where adjustments such as subject, start time, end time, last modified time
            and etc are set. This dicates the response.
    """

    return  """<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><soap:Header><t:RequestServerVersion Version="Exchange2007_SP1"/></soap:Header><soap:Body><FindItem Traversal="Shallow" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">%s<CalendarView StartDate="%s" EndDate="%s"/><ParentFolderIds><t:DistinguishedFolderId Id="calendar"/></ParentFolderIds></FindItem></soap:Body></soap:Envelope>""" % (shape, start, end) 


def default_shape():
    """ Returns an xml string of the default shape of a calendar item request
    """
    return """<ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties><t:FieldURI FieldURI="calendar:Start"/><t:FieldURI FieldURI="calendar:End"/><t:FieldURI FieldURI="item:LastModifiedTime"/><t:FieldURI FieldURI="item:IsUnmodified"/><t:FieldURI FieldURI="item:Subject"/></t:AdditionalProperties></ItemShape>"""


def try_url(url, path, port):
    """ Tries to open the url on the given port
        returns either a tuple of ('Redirect', LOCATION_TO_REDIRECT)
                                or ('Response', RESPONSE)
                                or ('FAILED', error_string) (if it failed)
    """ 

    print "Creating connection to %s%s on port %s" % (str(url), str(path), str(port))
    try:
        test = httplib.HTTPConnection(url, port)
        test.request('GET', path)
        
        response = test.getresponse()
        # see if there's a location, handle appropriately
        print "In attempt, found response: %s" % (response.read())
        
        location = response.getheader("Location")
        if (location != None):
        	return ("Redirect", location)
        
        else:
        	return ("Response", response.read())
    except:
        # Something broke, maybe it couldn't connect over ssl, or something else
        return ("FAILED", sys.exc_info()[1])


def try_ssl(url, path):
    """ Tries to open an ssl connection on the url with the given path """
    return try_url(url, path, 443)


def try_get(url, path):
    """ Tries to open a connection on the url with the given path """
    return try_url(url, path, 80)

def try_autodiscover(url, email, pw):
    """ 'Tries' a url according to the spec for calling autodiscover.
        Sends a 'GET' request to the url, if it succeeds, sends an HTTPS
        request, if valid, sends an authenticated 'POST' request.
        If the value of the response is a redirect, recurse on the new url.
        
        Returns the Autodiscover response if it gets one, otherwise returns 
        false.
    """
    auto_path = "/autodiscover/autodiscover.xml"

    def handle_attempt(attempt):
        print "Handle attempt: ", attempt
        result = attempt[0]
        if (result == "Redirect"):
            return auto_redirect(attempt[1], email, pw, 0)
        elif (result == "Response"):
            return attempt[1]
        else:
            # verify the try____ call is working
    	    assert result == "FAILED" 
    	    return False

    
    try1 = try_get(url, auto_path)
    res1 = handle_attempt(try1)

    if not res1:
    	print "try_get %s failed with %s" % (url, str(res1))
        # the request failed, so just bail out

    elif isinstance(res1, xml.dom.minidom.Document):
        return res1
    
    # Might have been a redirect to an https address, so try again
    print "trying ssl %s" % str(url)
    try2 = try_ssl(url, auto_path)
    res2 = handle_attempt(try2)

    if not res2:
        print "failed to https connect to %s" % url
    elif isinstance(res2, xml.dom.minidom.Document):
        return res2
    else:
    	print "https connect to %s returned %s" % (url, str(res2))
        # Send the authenticated POST request and see what happends
        try_post = autoredirect(url, email, pw, 0) 
        return try_post


def auto_redirect(url, email, pw, num):
    """ If the result of trying an autodiscover endpoing is a redirect,
        assume its a valid https endpoint, then send it xml. Don't follow more
        than 10 redirects according to the autodiscover client spec
    """

    if (num > 10):
    	return False
    else:
        try:
            try_post = send_xml(url, autodiscover_xml(email),
                    default_opener(), basic_auth_header(email, pw))

            # try_post might be an HTTPMessage telling us to redirect, so
            # check if thats the case. If so, redo the request. This is because
            # the httplib doesn't repost the xml data, so do it manually AND keep
            # track of the number of redirects at the same time
            if (isinstance(try_post, httplib.HTTPMessage)):
            	location = try_post.getheader('Location')
                if (location == None):
                	return False
                else:
                	return auto_redirect(location, email, pw, num + 1)
            else:
            	body = str(try_post.read())
                print "autoredirected xml send -> %s" % body
                return parseString(body)
        except:
            print "autoredicted error: %s" % sys.exc_info()[1]


def send_xml(url, xml, opener, auth):
    """ Sends the given xml to the url , with the given opener.
        If there is an authheader, it will be added to the request.
        NOTE :: encodes the xml to utf-8, and handles the default headers 
        Returns the response object directly.
    """

    xml_to_send = unicode(xml, 'utf-8')

    headers =  { "User-Agent": 'bacon', 
                "Content-Length": str(len(xml_to_send)),
                "Content-Type": "text/xml; charset=utf-8" }

    if auth:
    	headers['Authorization'] = auth

    print headers

    request = urllib2.Request(url=url, data=xml_to_send, headers=headers)

    print str(request)

    return opener.open(request)
    

def basic_auth_header(user, password):
    """ Returns a header for basic auth of the username and password.
        Base64 encodes user:password
    """
    
    base64userpass = base64.encodestring("%s:%s" % (user, password))[:-1]

    return "Basic %s" % base64userpass


def ntlm_opener(url, user, password):
    """ Returns a urllib2.OpenDirector for the given url, that requires NTLM authentication. """

    pass_mangr = urllib2.HTTPPasswordMgrWithDefaultRealm()
    pass_mangr.add_password(None, url, user, password)
    
    auth_NTLM = HTTPNtlmAuthHandler.HTTPNtlmAuthHandler(pass_mangr)

    return urllib2.build_opener(auth_NTLM)


class RedirectHandler(urllib2.HTTPRedirectHandler):
        def http_error_301(self, req, fp, code, msg, headers):
                result = urllib2.HTTPRedirectHandler.http_error_301( 
                        self, req, fp, code, msg, headers)
                result.status = code
                raise Exception("Permanent Redirect: %s" % 301)

        def http_error_302(self, req, fp, code, msg, headers):
                result = urllib2.HTTPRedirectHandler.http_error_302(
                        self, req, fp, code, msg, headers) 
                result.status = code
                print headers
                return headers


def default_opener():
    """ The default opener to use for sending xml requests """
    return urllib2.build_opener(RedirectHandler)


def get_domain(email):
    """ Return the domain name for a given email """
    
    # make a split in the form (<username>, '@', <domain>)
    split = email.partition('@')
    assert split[1] == '@'
    
    return split[2]

def get_domain_uname(email):
    """ Returns a Domain\username string for a given email """
    
    # make a split, similiary to the get_domain function
    split = email.partition('@')
    assert split[1] == '@'

    return "%s\%s" % (split[2], split[0])


def autodiscover(email, password):
    """ Follows the steps of the autodiscover process to retrieve xml for a
        valid autodiscvoer response. If its unable to find one, it returns False.
        NOTE:: the api to autoredirect is passed 0, because according ot the
        autodiscove spec we shouldn't follow more than 10 redirects. 

        TODO:: have auto_redirect check certificates to ssl connections 
        and not blindly send username:password's ...
    """
    domain = get_domain(email)

    url = "https://%s/autodiscover/autodiscover.xml" % domain
    attempt = auto_redirect(url, email, password, 0)

    if attempt:
        return attempt

    url = "https://autodiscover.%s/autodiscover/autodiscover.xml" % domain
    attempt = auto_redirect(url, email, password, 0)

    if attempt:
    	return attempt

    # if all else fails, use try_autodiscover, which will send unauthenticated get-requests and
    # hopefully succeed in getting a redirect location to start the auto_redirect process

    get_url = "autodiscover.%s" % domain
    attempt = try_autodiscover(get_url, email, password)

    if attempt:
    	return attempt

    # if there's control-flow here, it failed
    return False


def get_ews_from_xml(xml_data):
    """ Assumes this is valid autodiscover response xml, and tries to find the ews
        url that should be in @Protocol / @Type=EXPR / @EwsUrl
    """

    protocols = xml_data.getElementsByTagName('Protocol')
    for proto in protocols:
        types = proto.getElementsByTagName('Type')
        if (len(types) > 0):
        	proto_type = types[0].firstChild.data
        	print "protocol: %s has proto_type: %s" % (proto.toxml(), proto_type)
        	
        	if (proto_type == "EXPR"):
        		return proto.getElementsByTagName('EwsUrl')[0].firstChild.data
        	else:
        		continue
        else:
        	return False
    # if something hasn't been returned by here, there's a problem
    return False


def calendar_items(email, password, start, end):
    """ Tries to use autodiscover to retrieve calendar items.
        Expects to be passed start and end as valid utc strings.
        Returns valid xml or false
    """

    auto_xml = autodiscover(email, password)

    print auto_xml
    
    if not auto_xml:
    	return False
    
    # check that the xml from auto_discover is valid
    assert len(auto_xml.getElementsByTagName('Autodiscover')) == 1

    ews_url = get_ews_from_xml(auto_xml)
    print ews_url

    if not ews_url:
    	return False

    cal_xml = calendaritem_xml(default_shape(), start, end)
    print cal_xml

    try:
        try_cal = send_xml(ews_url, cal_xml, default_opener(), basic_auth_header(email, password))

        if try_cal:
        	cal_item_xml = parseString(try_cal.read())
        	print cal_item_xml.toprettyxml()
        	print try_cal.headers
        	return cal_item_xml
        else:
        	return False

    except urllib2.HTTPError, e:
        # check to see if its a 401 error, if so, there's a chance this https endpoint 
        # requires ntlm authentication. If thats the case, try again using the ntlm package
        
        if (401 == e.code):
        	print "HTTPError401, trying ntlm"
        	return ntlm_calendar_items(ews_url, email, password, cal_xml)
        else:
            print "Error:%s" % str(e)
            return False
    except:
        print "Unknown error: %s" % str(sys.exc_info()[1])
        return False

def ntlm_calendar_items(ews_url, email, pw, cal_xml):
    """ Tries to retrieve calendar item data with the given calitem request xml,
        at the destination ews_url, with the given email and pw.
        Determines the domain through the email and creates an ntlm opener.
    """

    # for ntlm we need to create a Domain\username from the email
    domuname = get_domain_uname(email)
    print domuname

    opener = ntlm_opener(ews_url, domuname, pw)
    print "made ntlm opener"

    print cal_xml

    try:
        # We don't need to add the basic-auth-header because the ntlm opener will
        # handle authorization the connection
        try_cal = send_xml(ews_url, cal_xml, opener, None)

        if try_cal:
        	cal_item_xml = parseString(try_cal.read())
        	print cal_item_xml.toprettyxml()
        	return cal_item_xml
        else:
        	return False

    except:
        print "Error: %s" % str(sys.exc_info()[1])
        return False

def get_cal_items(email, pw):
    return calendar_items(email, pw, START, END)
