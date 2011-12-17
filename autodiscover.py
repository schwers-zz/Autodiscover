import httplib, urllib2
import sys
import base64
import logging
import xml # for isinstance calls
from xml.dom.minidom import parseString
from ntlm import HTTPNtlmAuthHandler
# http://code.google.com/p/python-ntlm/

# This code is used for communicating with exchange. With an email and password
# it should be possible to follow the autodiscover protocol, and find an EWS url.
# EWS is Exchange Web Services, and can be queried for specifics about a specific
# user's calendar events, or even the mutual availability of users

# URL for implenting an Autodiscover-Client:
# http://msdn.microsoft.com/en-us/library/ee332364(v=EXCHG.140).aspx


# Set for Debug output
logger = logging.getLogger('autodiscover')
hdlr = logging.FileHandler('./auto.log')
formatter = logging.Formatter('%(asctime)s %(funcName)s %(levelname)s %(message)s')
hdlr.setFormatter(formatter)
# Change the log output level
hdlr.setLevel(logging.DEBUG)
logger.addHandler(hdlr)
logger.setLevel(logging.DEBUG)

# Define Helper Classes

class RedirectHandler(urllib2.HTTPRedirectHandler):
    """ urllib2's default URLOpener follows redirects. It does not however
        repost the xml. To deal with this, when a redirect is found,
        return the headers and let the calling function deal with things.
    """

    def handle_redirect(self, req, fp, code, msg, headers):
        logger.debug("http redirect headers found: %s" % headers)
        return headers

        return headers

    http_error_301 = handle_redirect
    http_error_302 = handle_redirect
    

class AutodiscoverError(Exception):
    def __init__(self, message):
        self.message = message


# Define Helper Functions

def xml_string_clean(xml_string):
    """ Cleaner for xml strings generated in the code. """

    xml_string = xml_string.replace('\n', '')
    xml_string = xml_string.replace('    ', '')

    return xml_string


def autodiscover_xml(email):
    """ Expects to have a string thats a valid email address for the intended exchange user
        Returns an xml string used to query the autodiscover servcie.
    """ 

    auto_xml = """<?xml version="1.0" encoding="utf-8"?>
    <Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
    <Request><EMailAddress>%(email)s</EMailAddress><AcceptableResponseSchema>
    http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a
    </AcceptableResponseSchema></Request></Autodiscover>""" % dict(email=email)

    return xml_string_clean(auto_xml)


def calendaritem_xml(shape, start, end):
    """ start & end -- valid utc strings designating a time range in which to query for 
            calendar events. 
        shpae -- xml string for the calendaritem shape according to the finditem spec, 
            this is where adjustments such as subject, start time, end time, last modified time
            and etc are set. This dicates the response.
    """

    cal_xml = """<?xml version="1.0" encoding="utf-8"?><soap:Envelope 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <soap:Header><t:RequestServerVersion Version="Exchange2007_SP1"/></soap:Header><soap:Body>
    <FindItem Traversal="Shallow" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
    %(shape)s<CalendarView StartDate="%(start)s" EndDate="%(end)s"/><ParentFolderIds>
    <t:DistinguishedFolderId Id="calendar"/>
    </ParentFolderIds></FindItem></soap:Body></soap:Envelope>""" % dict(shape=shape, start=start, end=end)

    return xml_string_clean(cal_xml)


def item_properties():
    """ Returns an xml string of the default shape of a calendar item request.
        A shape is set of fields returned with each calendar item.
        NOTE :: start here for further shape properties
            http://msdn.microsoft.com/en-us/library/aa564515(EXCHG.140).aspx
    """

    shape_xml = """<ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>
    <t:FieldURI FieldURI="calendar:Start"/>
    <t:FieldURI FieldURI="calendar:End"/>
    <t:FieldURI FieldURI="item:Subject"/>
    <t:FieldURI FieldURI="item:LastModifiedTime"/>
    </t:AdditionalProperties></ItemShape>"""

    return xml_string_clean(shape_xml)


def autodiscover_get_method(url, email, pw):
    """ The GET request part of the autodiscover client protocol. 
        Try sending a GET request to autodiscover.domain/autodiscover/autodiscover.xml

        If there's a redirect, go follow it with the autoredirect code.
    """

    path = "/autodiscover/autodiscover.xml"
    url = "autodiscover." + url

    test = httplib.HTTPConnection(url, 80)
    test.request('GET', path)

    response = test.getresponse()
    location = response.getheader("Location")

    # If there's a redirect, go follow it, otherwise, fail
    if location is None:
        raise AutodiscoverError("autodiscover_get_method failed for %s" % url)

    return auto_redirect(location, email, pw)
    

def auto_redirect(url, email, pw):
    """ If the result of trying an autodiscover endpoing is a redirect,
        assume its a valid https endpoint, then send it xml. Don't follow more
        than 10 redirects according to the autodiscover client spec
    """

    for i in range(10):
        logger.debug("Redirecting %(user)s to \
                autodiscover potential: %(url)s" % dict(user=email, url=url))
        try:
            response = send_xml(url, email, pw, autodiscover_xml(email))

            # response might be an HTTPMessage telling us to redirect, so
            # check if thats the case. If so, redo the request
            if isinstance(response, httplib.HTTPMessage):
                location = response.getheader('Location')
                if location is None:
                    # The http message headers should come from RedirectHandler
                    # therefore, there should be a 'Location' header
                    raise AutodiscoverError("Recieved HTTPMessage without redirect header")
                else:
                    # otherwise continue
                    url = location
                    continue
                    
            else:
                # A response was recieved, try to parse it and return it
                body = str(response.read())
                logger.debug("autoredirect xml sending: %s" % body)
                return parseString(body)

        except AutodiscoverError, e:
            logger.debug("Autodiscover exception caught in auto_redirect: %s" % str(e))
            raise e
    
    # More than 10 requests were made, or some other error occured
    error_str = "auto_redirect failed to autodiscover $(url)s \
            for %(email)s:%(pw)s" % dict(url=url, email=email, pw=pw)

    raise AutodiscoverError(error_str)


def send_xml(url, email, pw, body_xml):
    """ Sends the given xml to the url , with the given opener.
        If there is an authheader, it will be added to the request.
        NOTE :: encodes the xml to utf-8, and handles the default headers 
        Returns the response object directly.
    """

    logger.debug("send_xml: %(xml)s" % dict(xml=send_xml))

    body_xml_send = body_xml.encode("utf-8")
     
    # Authentication headers have to be added, because urllib2 first sends an
    # unauthenticated request, so things break prematurely
    b64userpw = base64.encodestring("%s:%s" % (email, pw)).strip()

    headers = {"Content-Type": "text/xml; charset=utf-8",
            "Authorization": "Basic %s" % b64userpw }

    request = urllib2.Request(url=url, data=body_xml_send, headers=headers)

    logger.debug("Sending: %s" % str(request))

    # try opening it with a default opener
    opener = default_opener(url, email, pw)
    try:
        return opener.open(request)
        
    # If that didn't work, and was a 401, it might require NTLM authentication
    except urllib2.HTTPError, e:
        if 401 != e.code:
            logger.warn("send_xml encountered unexpected \
                    httperror %(error)s when posting to url:%(url)s" % dict(error=e, url=url))
            raise e
        else:
            opener = ntlm_opener(url, email, pw)
            return opener.open(request)
    
    # It might have been another http related error, like the url being invalid
    except urllib2.URLError, e:
        raise AutodiscoverError("URLerror caught in send_xml: %s" % e)


def default_opener(url, email, pw):
    """ The default opener to use for sending xml requests."""

    return urllib2.build_opener(RedirectHandler)


def ntlm_opener(url, email, password):
    """ Returns a urllib2.OpenDirector for the given url, that requires NTLM authentication. """

    # the ntlm library requires Domain\user format
    user = get_domain_uname(email)

    pass_mangr = urllib2.HTTPPasswordMgrWithDefaultRealm()
    pass_mangr.add_password(None, url, user, password)
    
    auth_NTLM = HTTPNtlmAuthHandler.HTTPNtlmAuthHandler(pass_mangr)

    return urllib2.build_opener(auth_NTLM)


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

    # Step 1. Try appending to the domain and using Authenticated Post Requests
    url = "https://%s/autodiscover/autodiscover.xml" % domain
    try:
        return auto_redirect(url, email, password)

    except AutodiscoverError, e:
        logger.debug("Failed attempt to autodiscover via %s, trying Step 2" % url)

    # Step 2. Try adding autodiscover. to the url
    url = "https://autodiscover.%s/autodiscover/autodiscover.xml" % domain
    try:
        return auto_redirect(url, email, password)

    except AutodiscoverError, e:
        logger.debug("Failed attempt to autodiscover via %s, trying Step 3" % url)

    # Step 3. Try sending an un-authenticated GET request to the previous url
    # IF There's a redirect, autodiscover_get_method will follow this
    # Note :: autodiscover_get_method will fix the domain to what it needs to be
    return autodiscover_get_method(domain, email, password)


def get_ews_from_xml(xml_data):
    """ Assumes this is valid autodiscover response xml, and tries to find the ews
        url that should be in @Protocol / @Type=EXPR / @EwsUrl
    """

    protocols = xml_data.getElementsByTagName('Protocol')
    for proto in protocols:
        types = proto.getElementsByTagName('Type')
        if types:
            proto_type = types[0].firstChild.data
            if "EXPR" == proto_type:
                return proto.getElementsByTagName('EwsUrl')[0].firstChild.data

    # nothing returned, raise an exception
    raise AutodiscoverError("Coudn't retrieve ews url from autodiscove response")


def calendar_items(email, password, start, end):
    """ Tries to use autodiscover to retrieve calendar items.
        Expects to be passed start and end as valid utc strings.
        Returns valid xml or false
    """

    auto_xml = autodiscover(email, password)
    
    # check that the xml from auto_discover is valid
    assert len(auto_xml.getElementsByTagName('Autodiscover')) == 1

    ews_url = get_ews_from_xml(auto_xml)
    logger.debug("Found ews url:%s" % ews_url)

    cal_xml = calendaritem_xml(item_properties(), start, end)

    cal_item_resp = send_xml(ews_url, email, pw, cal_xml)

    cal_item_xml = parseString(cal_item_resp.read())
    logger.debug("Found calendar item xml: %s" % cal_item_xml.toxml())

    return cal_item_xml


if __name__ == '__main__':
    if 2 > len(sys.argv):
        print "usage: autodiscover.py email password"
        sys.exit()

    # Default time range Useful for Testing
    START = "2010-12-12T00:00:00-08:00"
    END = "2011-12-19T00:00:00-08:00"
    
    email = sys.argv[1]
    pw = sys.argv[2]

    print calendar_items(email, pw, START, END).toprettyxml()
    sys.exit()
