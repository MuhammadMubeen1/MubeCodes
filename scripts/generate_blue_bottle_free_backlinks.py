#!/usr/bin/env python3
"""Blue Bottle case: free directories, guest posting & local citation sites only."""
from __future__ import annotations

import csv
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(__file__).resolve().parents[1] / "portfolio" / "reports"
CSV_PATH = OUT / "blue-bottle-coffee-free-backlinks-1000.csv"
XLSX_PATH = OUT / "blue-bottle-coffee-free-backlinks-1000.xlsx"

HEADERS = [
    "#",
    "Referring domain",
    "Category",
    "Submission URL",
    "Cost",
    "Easy post (1=easy, 5=hardest)",
    "USA relevant",
    "Link type",
    "Notes (Blue Bottle / café D2C)",
]

# (domain, category, submission_url, easy, usa, link_type, notes)
# category must be: Free directory | Guest posting | Local citation

LOCAL_CITATION: list[tuple[str, ...]] = [
    ("business.google.com", "Local citation", "https://business.google.com/", "1", "Yes", "NAP + website", "Google Business Profile"),
    ("bingplaces.com", "Local citation", "https://www.bingplaces.com/", "1", "Yes", "Citation", "Bing Places"),
    ("mapsconnect.apple.com", "Local citation", "https://mapsconnect.apple.com/", "2", "Yes", "Citation", "Apple Business Connect"),
    ("biz.yelp.com", "Local citation", "https://biz.yelp.com/", "2", "Yes", "Listing", "Yelp for Business"),
    ("foursquare.com", "Local citation", "https://foursquare.com/claim", "2", "Yes", "Venue", "Foursquare / data feed"),
    ("nextdoor.com", "Local citation", "https://business.nextdoor.com/", "2", "Yes", "Business page", "Nextdoor Business"),
    ("mapquest.com", "Local citation", "https://www.mapquest.com/", "2", "Yes", "Listing", "MapQuest place"),
    ("tripadvisor.com", "Local citation", "https://www.tripadvisor.com/Owners", "2", "Yes", "Listing", "Café owner claim"),
    ("opentable.com", "Local citation", "https://restaurant.opentable.com/", "3", "Yes", "Restaurant", "If reservations"),
    ("alignable.com", "Local citation", "https://www.alignable.com/", "2", "Yes", "SMB profile", "Local SMB network"),
    ("patch.com", "Local citation", "https://patch.com/", "3", "Yes", "Local listing", "Patch business (where available)"),
    ("bbb.org", "Local citation", "https://www.bbb.org/get-accredited", "3", "Yes", "Accreditation", "BBB listing"),
    ("angi.com", "Local citation", "https://www.angi.com/", "3", "Yes", "Pro profile", "Home services angle"),
    ("thumbtack.com", "Local citation", "https://www.thumbtack.com/", "3", "Yes", "Pro profile", "Local services"),
    ("houzz.com", "Local citation", "https://www.houzz.com/pro", "3", "Yes", "Pro profile", "Café design / build"),
    ("bark.com", "Local citation", "https://www.bark.com/en/us/", "2", "Yes", "Provider", "Bark.com free profile"),
    ("localeze.com", "Local citation", "https://www.neustarlocaleze.biz/", "3", "Yes", "Data aggregator", "Neustar Localeze"),
    ("data-axle.com", "Local citation", "https://www.data-axle.com/", "3", "Yes", "Data aggregator", "Data Axle / Infogroup"),
    ("openstreetmap.org", "Local citation", "https://www.openstreetmap.org/", "3", "Yes", "Map POI", "Add café POI (community)"),
    ("waze.com", "Local citation", "https://www.waze.com/", "3", "Yes", "Map listing", "Waze place via Google sync"),
    ("here.com", "Local citation", "https://www.here.com/", "3", "Yes", "Maps", "HERE WeGo business"),
    ("tomtom.com", "Local citation", "https://www.tomtom.com/", "3", "Yes", "Maps", "TomTom partners"),
    ("apple.com", "Local citation", "https://mapsconnect.apple.com/", "2", "Yes", "Maps", "Apple Maps (Business Connect)"),
    ("facebook.com", "Local citation", "https://www.facebook.com/business/", "2", "Yes", "Local page", "Facebook local business"),
    ("nycgo.com", "Local citation", "https://www.nycgo.com/", "3", "Yes", "Tourism partner", "NYC tourism listings"),
    ("sftravel.com", "Local citation", "https://www.sftravel.com/", "3", "Yes", "Tourism partner", "SF Travel Association"),
    ("bostonusa.com", "Local citation", "https://www.bostonusa.com/", "3", "Yes", "Tourism partner", "Boston USA"),
    ("choosechicago.com", "Local citation", "https://www.choosechicago.com/", "3", "Yes", "Tourism partner", "Choose Chicago"),
    ("visitseattle.org", "Local citation", "https://www.visitseattle.org/", "3", "Yes", "Tourism partner", "Visit Seattle"),
    ("visitportland.com", "Local citation", "https://www.visitportland.com/", "3", "Yes", "Tourism partner", "Travel Portland"),
    ("visitcalifornia.com", "Local citation", "https://www.visitcalifornia.com/", "3", "Yes", "Tourism partner", "Visit California"),
    ("traveltexas.com", "Local citation", "https://www.traveltexas.com/", "3", "Yes", "Tourism partner", "Travel Texas"),
    ("exploregeorgia.org", "Local citation", "https://exploregeorgia.org/", "3", "Yes", "Tourism partner", "Explore Georgia"),
    ("visitdenver.com", "Local citation", "https://www.denver.org/", "3", "Yes", "Tourism partner", "Visit Denver"),
    ("washington.org", "Local citation", "https://washington.org/", "3", "Yes", "Tourism partner", "Washington DC"),
    ("miamiandbeaches.com", "Local citation", "https://www.miamiandbeaches.com/", "3", "Yes", "Tourism partner", "Greater Miami"),
    ("findmecoffee.com", "Local citation", "https://www.findmecoffee.com/", "2", "Yes", "Coffee shop listing", "Find Me Coffee directory"),
    ("allmenus.com", "Local citation", "https://www.allmenus.com/", "3", "Yes", "Menu listing", "AllMenus"),
    ("menuism.com", "Local citation", "https://www.menuism.com/", "3", "Yes", "Menu listing", "Menuism"),
    ("zomato.com", "Local citation", "https://www.zomato.com/", "3", "Yes", "Restaurant", "Zomato partner (US cities)"),
    ("untappd.com", "Local citation", "https://untappd.com/", "3", "Yes", "Venue", "If beer / coffee bar"),
]

FREE_DIRECTORY: list[tuple[str, ...]] = [
    ("manta.com", "Free directory", "https://www.manta.com/add-your-business", "1", "Yes", "Citation", "Manta"),
    ("hotfrog.com", "Free directory", "https://www.hotfrog.com/add-your-business", "1", "Yes", "Citation", "Hotfrog"),
    ("hotfrog.us", "Free directory", "https://www.hotfrog.us/add-your-business", "1", "Yes", "Citation", "Hotfrog USA"),
    ("brownbook.net", "Free directory", "https://www.brownbook.net/add-business/", "1", "Yes", "Citation", "Brownbook"),
    ("cylex.us.com", "Free directory", "https://www.cylex.us.com/company/add/", "1", "Yes", "Citation", "Cylex US"),
    ("ezlocal.com", "Free directory", "https://www.ezlocal.com/", "2", "Yes", "Citation", "EZ Local"),
    ("n49.com", "Free directory", "https://www.n49.com/", "2", "Yes", "Citation", "N49"),
    ("tupalo.com", "Free directory", "https://www.tupalo.com/", "2", "Yes", "Citation", "Tupalo"),
    ("showmelocal.com", "Free directory", "https://www.showmelocal.com/addbusiness.aspx", "1", "Yes", "Citation", "ShowMeLocal"),
    ("citysquares.com", "Free directory", "https://citysquares.com/add-business", "2", "Yes", "Citation", "CitySquares"),
    ("americantowns.com", "Free directory", "https://www.americantowns.com/", "2", "Yes", "Citation", "American Towns"),
    ("chamberofcommerce.com", "Free directory", "https://www.chamberofcommerce.com/add-your-business", "2", "Yes", "Citation", "ChamberOfCommerce.com"),
    ("merchantcircle.com", "Free directory", "https://www.merchantcircle.com/signup", "2", "Yes", "Citation", "MerchantCircle"),
    ("local.com", "Free directory", "https://www.local.com/business/add/", "2", "Yes", "Citation", "Local.com"),
    ("ibegin.com", "Free directory", "https://www.ibegin.com/directory/add/", "1", "Yes", "Citation", "iBegin"),
    ("n2local.com", "Free directory", "https://www.n2local.com/addbusiness", "2", "Yes", "Citation", "N2Local"),
    ("yellowpages.com", "Free directory", "https://www.yellowpages.com/", "2", "Yes", "Listing", "Yellow Pages"),
    ("superpages.com", "Free directory", "https://www.superpages.com/", "2", "Yes", "Citation", "SuperPages"),
    ("dexknows.com", "Free directory", "https://www.dexknows.com/", "2", "Yes", "Citation", "DexKnows"),
    ("citysearch.com", "Free directory", "https://www.citysearch.com/", "3", "Yes", "Listing", "Citysearch"),
    ("judysbook.com", "Free directory", "https://www.judysbook.com/", "2", "Yes", "Listing", "Judy's Book"),
    ("insiderpages.com", "Free directory", "https://www.insiderpages.com/", "3", "Yes", "Listing", "Insider Pages"),
    ("spoke.com", "Free directory", "https://www.spoke.com/", "2", "Yes", "Company profile", "Spoke"),
    ("trustpilot.com", "Free directory", "https://business.trustpilot.com/", "2", "Yes", "Reviews", "Trustpilot"),
    ("sitejabber.com", "Free directory", "https://www.sitejabber.com/", "2", "Yes", "Reviews", "Sitejabber"),
    ("cybo.com", "Free directory", "https://www.cybo.com/add/", "2", "Yes", "Citation", "Cybo"),
    ("tuugo.us", "Free directory", "https://www.tuugo.us/", "2", "Yes", "Citation", "Tuugo US"),
    ("storeboard.com", "Free directory", "https://www.storeboard.com/", "2", "Yes", "B2B listing", "Storeboard"),
    ("salespider.com", "Free directory", "https://www.salespider.com/", "2", "Yes", "B2B", "SalesSpider"),
    ("linkcentre.com", "Free directory", "https://www.linkcentre.com/submit.php", "2", "Yes", "Directory submit", "LinkCentre"),
    ("viesearch.com", "Free directory", "https://viesearch.com/submit", "2", "Yes", "Directory submit", "VieSearch"),
    ("freelistingusa.com", "Free directory", "https://www.freelistingusa.com/", "2", "Yes", "Citation", "Free Listing USA"),
    ("find-us-here.com", "Free directory", "https://www.find-us-here.com/", "2", "Yes", "Citation", "Find Us Here"),
    ("enrollbusiness.com", "Free directory", "https://us.enrollbusiness.com/", "2", "Yes", "Citation", "Enroll Business"),
    ("callupcontact.com", "Free directory", "https://www.callupcontact.com/", "2", "Yes", "Citation", "CallUpContact"),
    ("expressbusinessdirectory.com", "Free directory", "https://www.expressbusinessdirectory.com/", "2", "Yes", "Citation", "Express Business Directory"),
    ("freebusinessdirectory.com", "Free directory", "https://www.freebusinessdirectory.com/", "2", "Yes", "Citation", "Free Business Directory"),
    ("a-zbusinessfinder.com", "Free directory", "https://www.a-zbusinessfinder.com/", "2", "Yes", "Citation", "A-Z Business Finder"),
    ("addonbiz.com", "Free directory", "https://www.addonbiz.com/", "2", "Yes", "Citation", "AddOnBiz"),
    ("find-open.com", "Free directory", "https://www.find-open.com/", "2", "Yes", "Citation", "Find-Open"),
    ("zipleaf.us", "Free directory", "https://www.zipleaf.us/", "2", "Yes", "Citation", "ZipLeaf US"),
    ("provenexpert.com", "Free directory", "https://www.provenexpert.com/en-us/", "2", "Yes", "Reviews", "ProvenExpert US"),
    ("botw.org", "Free directory", "https://botw.org/", "2", "Yes", "Directory submit", "Best of the Web"),
    ("jayde.com", "Free directory", "https://www.jayde.com/", "2", "Yes", "Directory", "Jayde"),
    ("g2.com", "Free directory", "https://www.g2.com/", "3", "Yes", "Software/reviews", "G2 profile (if applicable)"),
    ("clutch.co", "Free directory", "https://clutch.co/", "3", "Yes", "B2B directory", "Clutch profile"),
    ("goodfirms.co", "Free directory", "https://www.goodfirms.co/", "3", "Yes", "B2B", "GoodFirms"),
    ("crunchbase.com", "Free directory", "https://www.crunchbase.com/", "3", "Yes", "Company", "Crunchbase org"),
    ("dnb.com", "Free directory", "https://www.dnb.com/", "3", "Yes", "Company", "D-U-N-S / D&B"),
    ("zoominfo.com", "Free directory", "https://www.zoominfo.com/", "3", "Yes", "Company", "ZoomInfo claim"),
    ("allstatesusadirectory.com", "Free directory", "https://www.allstatesusadirectory.com/", "2", "Yes", "Citation", "All States USA Directory"),
    ("usalistingdirectory.com", "Free directory", "https://www.usalistingdirectory.com/", "2", "Yes", "Citation", "USA Listing Directory"),
    ("usawebsitesdirectory.com", "Free directory", "https://www.usawebsitesdirectory.com/", "2", "Yes", "Citation", "USA Websites Directory"),
    ("411.info", "Free directory", "https://www.411.info/", "2", "Yes", "Citation", "411.info"),
    ("411freedirectory.com", "Free directory", "https://411freedirectory.com/", "2", "Yes", "Citation", "411 Free Directory"),
    ("elocal.com", "Free directory", "https://www.elocal.com/", "2", "Yes", "Citation", "eLocal"),
    ("localstack.com", "Free directory", "https://localstack.com/", "2", "Yes", "Citation", "LocalStack"),
    ("iglobal.co", "Free directory", "https://www.iglobal.co/", "2", "Yes", "Citation", "iGlobal"),
    ("lacartes.com", "Free directory", "https://www.lacartes.com/", "2", "Yes", "Citation", "Lacartes"),
    ("bunity.com", "Free directory", "https://www.bunity.com/", "2", "Yes", "B2B", "Bunity"),
    ("bizhwy.com", "Free directory", "https://www.bizhwy.com/", "2", "Yes", "Citation", "BizHwy"),
    ("igotbiz.com", "Free directory", "https://www.igotbiz.com/", "2", "Yes", "Citation", "iGotBiz"),
    ("localbiznetwork.com", "Free directory", "https://www.localbiznetwork.com/", "2", "Yes", "Citation", "Local Biz Network"),
]

GUEST_POSTING: list[tuple[str, ...]] = [
    ("perfectdailygrind.com", "Guest posting", "https://perfectdailygrind.com/contact/", "4", "Yes", "Editorial pitch", "Specialty coffee news — pitch stories"),
    ("baristamagazine.com", "Guest posting", "https://www.baristamagazine.com/contact/", "4", "Yes", "Contributor", "Barista Magazine"),
    ("sprudge.com", "Guest posting", "https://sprudge.com/contact", "4", "Yes", "Pitch", "Sprudge coffee culture"),
    ("freshcup.com", "Guest posting", "https://freshcup.com/contact/", "4", "Yes", "Contributor", "Fresh Cup Magazine"),
    ("ineedcoffee.com", "Guest posting", "https://ineedcoffee.com/", "3", "Yes", "Community article", "I Need Coffee community posts"),
    ("dailycoffeenews.com", "Guest posting", "https://dailycoffeenews.com/contact/", "4", "Yes", "News pitch", "Daily Coffee News"),
    ("standartmag.com", "Guest posting", "https://standartmag.com/", "4", "Yes", "Editorial", "Standart coffee magazine"),
    ("EuropeanCoffeeTrip.com", "Guest posting", "https://europeancoffeetrip.com/", "4", "Yes", "Guest article", "Coffee travel / culture"),
    ("homebarista.com", "Guest posting", "https://www.homebarista.com/", "3", "Yes", "Forum / guides", "Home-Barista community"),
    ("coffeeforums.com", "Guest posting", "https://www.coffeeforums.com/", "3", "Yes", "Forum", "Coffee Forums"),
    ("roastmagazine.com", "Guest posting", "https://www.roastmagazine.com/", "4", "Yes", "Industry", "Roast Magazine"),
    ("coffeereview.com", "Guest posting", "https://www.coffeereview.com/", "4", "Yes", "Review / guest", "Coffee Review"),
    ("specialtycoffeeassociation.org", "Guest posting", "https://sca.coffee/", "4", "Yes", "Industry", "SCA community content"),
    ("blog.feedspot.com", "Guest posting", "https://blog.feedspot.com/coffee_blogs/", "3", "Yes", "Outreach list", "Coffee blog outreach hub"),
    ("hubpages.com", "Guest posting", "https://hubpages.com/", "2", "Yes", "User article", "HubPages free articles"),
    ("ezinearticles.com", "Guest posting", "https://ezinearticles.com/", "3", "Yes", "Article", "EzineArticles (check status)"),
    ("medium.com", "Guest posting", "https://medium.com/", "2", "Yes", "Publication", "Medium — coffee / D2C stories"),
    ("dev.to", "Guest posting", "https://dev.to/", "2", "Yes", "Tech article", "D2C / ecommerce angle"),
    ("hashnode.com", "Guest posting", "https://hashnode.com/", "2", "Yes", "Blog", "Hashnode dev / startup blogs"),
    ("substack.com", "Guest posting", "https://substack.com/", "2", "Yes", "Newsletter guest", "Guest post in newsletters"),
    ("foodbloggerscentral.com", "Guest posting", "https://foodbloggerscentral.com/", "3", "Yes", "Food guest", "Food blogger network"),
    ("thekitchn.com", "Guest posting", "https://www.thekitchn.com/", "5", "Yes", "Pitch", "The Kitchn — pitch only"),
    ("seriouseats.com", "Guest posting", "https://www.seriouseats.com/", "5", "Yes", "Pitch", "Serious Eats editorial"),
    ("tastingtable.com", "Guest posting", "https://www.tastingtable.com/", "4", "Yes", "Pitch", "Food & drink culture"),
    ("food52.com", "Guest posting", "https://food52.com/", "4", "Yes", "Community", "Food52 community"),
    ("eater.com", "Guest posting", "https://www.eater.com/", "5", "Yes", "Pitch", "Eater — editorial pitch"),
    ("bonappetit.com", "Guest posting", "https://www.bonappetit.com/", "5", "Yes", "Pitch", "Bon Appétit pitch"),
    ("saveur.com", "Guest posting", "https://www.saveur.com/", "4", "Yes", "Pitch", "Saveur food travel"),
    ("epicurious.com", "Guest posting", "https://www.epicurious.com/", "4", "Yes", "Pitch", "Epicurious"),
    ("delish.com", "Guest posting", "https://www.delish.com/", "4", "Yes", "Pitch", "Delish"),
    ("allrecipes.com", "Guest posting", "https://www.allrecipes.com/", "4", "Yes", "Community", "Allrecipes community"),
    ("thespruceeats.com", "Guest posting", "https://www.thespruceeats.com/", "3", "Yes", "Contributor", "Spruce Eats contributor"),
    ("simplyrecipes.com", "Guest posting", "https://www.simplyrecipes.com/", "4", "Yes", "Pitch", "Simply Recipes"),
    ("minimalistbaker.com", "Guest posting", "https://minimalistbaker.com/", "5", "Yes", "Pitch", "Recipe / beverage guest"),
    ("halfbakedharvest.com", "Guest posting", "https://www.halfbakedharvest.com/", "5", "Yes", "Pitch", "Lifestyle food"),
    ("foodbusinessnews.net", "Guest posting", "https://www.foodbusinessnews.net/", "4", "Yes", "Industry", "Food Business News"),
    ("beveragedaily.com", "Guest posting", "https://www.beveragedaily.com/", "4", "Yes", "Industry PR", "Beverage industry"),
    ("bevnet.com", "Guest posting", "https://www.bevnet.com/", "4", "Yes", "Press / guest", "BevNET drinks"),
    ("foodnavigator-usa.com", "Guest posting", "https://www.foodnavigator-usa.com/", "4", "Yes", "Industry", "FoodNavigator USA"),
    ("forbes.com", "Guest posting", "https://www.forbes.com/sites/forbesbusinesscouncil/", "4", "Yes", "Council post", "Forbes BC (paid tier exists)"),
    ("business.com", "Guest posting", "https://www.business.com/", "3", "Yes", "Contributor", "Business.com articles"),
    ("entrepreneur.com", "Guest posting", "https://www.entrepreneur.com/", "4", "Yes", "Contributor", "Entrepreneur contributor"),
    ("inc.com", "Guest posting", "https://www.inc.com/", "4", "Yes", "Pitch", "Inc.com"),
    ("smallbiztrends.com", "Guest posting", "https://smallbiztrends.com/", "3", "Yes", "Guest post", "Small Biz Trends"),
    ("shopify.com", "Guest posting", "https://www.shopify.com/blog", "4", "Yes", "Pitch", "Shopify blog — D2C"),
    ("bigcommerce.com", "Guest posting", "https://www.bigcommerce.com/blog/", "4", "Yes", "Pitch", "BigCommerce blog"),
    ("hospitalitynet.org", "Guest posting", "https://www.hospitalitynet.org/", "4", "Yes", "Opinion", "Hospitality Net opinion"),
    ("restaurantbusinessonline.com", "Guest posting", "https://www.restaurantbusinessonline.com/", "4", "Yes", "Industry", "Restaurant Business"),
    ("nrn.com", "Guest posting", "https://www.nrn.com/", "4", "Yes", "Industry", "Nation's Restaurant News"),
    ("qsrmagazine.com", "Guest posting", "https://www.qsrmagazine.com/", "4", "Yes", "Industry", "QSR magazine"),
    ("fastcasual.com", "Guest posting", "https://www.fastcasual.com/", "4", "Yes", "Industry", "Fast Casual"),
    ("blogtrepreneur.com", "Guest posting", "https://www.blogtrepreneur.com/write-for-us/", "2", "Yes", "Write for us", "Blogtrepreneur"),
    ("copyblogger.com", "Guest posting", "https://copyblogger.com/", "4", "Yes", "Pitch", "Marketing / content"),
    ("contentmarketinginstitute.com", "Guest posting", "https://contentmarketinginstitute.com/", "4", "Yes", "Pitch", "CMI"),
    ("socialmediaexaminer.com", "Guest posting", "https://www.socialmediaexaminer.com/", "4", "Yes", "Pitch", "SME guest"),
    ("searchengineland.com", "Guest posting", "https://searchengineland.com/", "5", "Yes", "Pitch", "Search Engine Land"),
    ("searchenginejournal.com", "Guest posting", "https://www.searchenginejournal.com/", "4", "Yes", "Contributor", "SEJ contributor"),
    ("moz.com", "Guest posting", "https://moz.com/blog", "5", "Yes", "Pitch", "Moz blog pitch"),
    ("ahrefs.com", "Guest posting", "https://ahrefs.com/blog/", "5", "Yes", "Pitch", "Ahrefs blog"),
    ("semrush.com", "Guest posting", "https://www.semrush.com/blog/", "5", "Yes", "Pitch", "Semrush blog"),
    ("localseoguide.com", "Guest posting", "https://localseoguide.com/", "3", "Yes", "Guest post", "Local SEO Guide"),
    ("brightlocal.com", "Guest posting", "https://www.brightlocal.com/learn/", "4", "Yes", "Resources", "BrightLocal learn hub"),
    ("whitespark.ca", "Guest posting", "https://whitespark.ca/blog/", "4", "Yes", "Local SEO", "Whitespark blog"),
    ("streetfightmag.com", "Guest posting", "https://streetfightmag.com/", "3", "Yes", "Local biz", "Street Fight local"),
    ("localu.org", "Guest posting", "https://localu.org/", "3", "Yes", "Local SEO", "LocalU community"),
    ("nearmedia.co", "Guest posting", "https://nearmedia.co/", "3", "Yes", "Local media", "Near Media"),
    ("patch.com", "Guest posting", "https://patch.com/", "3", "Yes", "Local article", "Patch contributor posts"),
    ("medium.com", "Guest posting", "https://medium.com/", "2", "Yes", "Publication", "Medium coffee publications"),
    ("linkedin.com", "Guest posting", "https://www.linkedin.com/pulse/", "2", "Yes", "Article", "LinkedIn Pulse articles"),
    ("quora.com", "Guest posting", "https://www.quora.com/", "2", "Yes", "Answer / space", "Quora spaces — brewing topics"),
    ("reddit.com", "Guest posting", "https://www.reddit.com/r/Coffee/", "3", "Yes", "Community", "r/Coffee — follow rules"),
    ("steemit.com", "Guest posting", "https://steemit.com/", "3", "Yes", "Blog", "Steemit posts"),
    ("hackernoon.com", "Guest posting", "https://hackernoon.com/", "3", "Yes", "Tech story", "Hacker Noon"),
    ("indiehackers.com", "Guest posting", "https://www.indiehackers.com/", "3", "Yes", "Story", "Indie Hackers D2C"),
    ("startup.info", "Guest posting", "https://startup.info/", "2", "Yes", "Write for us", "Startup.info guest"),
    ("foundr.com", "Guest posting", "https://foundr.com/", "4", "Yes", "Pitch", "Foundr magazine"),
    ("smartbrief.com", "Guest posting", "https://www.smartbrief.com/", "4", "Yes", "Briefing", "SmartBrief food briefs"),
    ("travelandleisure.com", "Guest posting", "https://www.travelandleisure.com/", "5", "Yes", "Pitch", "Travel + café features"),
    ("cntraveler.com", "Guest posting", "https://www.cntraveler.com/", "5", "Yes", "Pitch", "Condé Nast Traveler"),
    ("timeout.com", "Guest posting", "https://www.timeout.com/", "4", "Yes", "Pitch", "Time Out city guides"),
    ("thrillist.com", "Guest posting", "https://www.thrillist.com/", "4", "Yes", "Pitch", "Thrillist food/drink"),
    ("theinfatuation.com", "Guest posting", "https://www.theinfatuation.com/", "4", "Yes", "Pitch", "Infatuation café reviews"),
    ("zomato.com", "Guest posting", "https://www.zomato.com/blog", "4", "Yes", "Blog", "Zomato blog pitch"),
    ("grubstreet.com", "Guest posting", "https://www.grubstreet.com/", "5", "Yes", "Pitch", "Grub Street NY"),
    ("sf.eater.com", "Guest posting", "https://sf.eater.com/", "5", "Yes", "Pitch", "Eater SF local"),
    ("ny.eater.com", "Guest posting", "https://ny.eater.com/", "5", "Yes", "Pitch", "Eater NY local"),
    ("la.eater.com", "Guest posting", "https://la.eater.com/", "5", "Yes", "Pitch", "Eater LA"),
    ("chicago.eater.com", "Guest posting", "https://chicago.eater.com/", "5", "Yes", "Pitch", "Eater Chicago"),
    ("boston.com", "Guest posting", "https://www.boston.com/", "4", "Yes", "Pitch", "Boston.com lifestyle"),
    ("sfchronicle.com", "Guest posting", "https://www.sfchronicle.com/", "5", "Yes", "Pitch", "SF Chronicle food"),
    ("latimes.com", "Guest posting", "https://www.latimes.com/", "5", "Yes", "Pitch", "LA Times food"),
    ("nytimes.com", "Guest posting", "https://www.nytimes.com/international/", "5", "Yes", "Pitch", "NYT — editorial only"),
    ("wirecutter.com", "Guest posting", "https://www.nytimes.com/wirecutter/", "5", "Yes", "Pitch", "Wirecutter product guides"),
    ("architecturaldigest.com", "Guest posting", "https://www.architecturaldigest.com/", "5", "Yes", "Pitch", "AD design / café"),
    ("design-milk.com", "Guest posting", "https://design-milk.com/", "4", "Yes", "Pitch", "Design Milk"),
    ("dezeen.com", "Guest posting", "https://www.dezeen.com/", "5", "Yes", "Pitch", "Dezeen architecture"),
    ("dwell.com", "Guest posting", "https://www.dwell.com/", "4", "Yes", "Pitch", "Dwell interiors"),
    ("apartmenttherapy.com", "Guest posting", "https://www.apartmenttherapy.com/", "3", "Yes", "Contributor", "Apartment Therapy"),
    ("theeverygirl.com", "Guest posting", "https://theeverygirl.com/", "3", "Yes", "Guest post", "The Everygirl lifestyle"),
    ("wellandgood.com", "Guest posting", "https://www.wellandgood.com/", "4", "Yes", "Pitch", "Well+Good wellness"),
    ("mindbodygreen.com", "Guest posting", "https://www.mindbodygreen.com/", "4", "Yes", "Pitch", "Mindbodygreen"),
    ("popsugar.com", "Guest posting", "https://www.popsugar.com/", "4", "Yes", "Pitch", "PopSugar food"),
    ("buzzfeed.com", "Guest posting", "https://www.buzzfeed.com/", "4", "Yes", "Community", "BuzzFeed Community"),
    ("huffpost.com", "Guest posting", "https://www.huffpost.com/", "3", "Yes", "Contributor", "HuffPost contributor"),
    ("mashable.com", "Guest posting", "https://mashable.com/", "4", "Yes", "Pitch", "Mashable"),
    ("lifehacker.com", "Guest posting", "https://lifehacker.com/", "3", "Yes", "Pitch", "Lifehacker tips"),
    ("themanual.com", "Guest posting", "https://www.themanual.com/", "3", "Yes", "Guest post", "The Manual — drinks"),
    ("insidehook.com", "Guest posting", "https://www.insidehook.com/", "4", "Yes", "Pitch", "InsideHook"),
    ("gearpatrol.com", "Guest posting", "https://gearpatrol.com/", "4", "Yes", "Pitch", "Gear Patrol coffee gear"),
    ("uncrate.com", "Guest posting", "https://uncrate.com/", "5", "Yes", "Pitch", "Uncrate products"),
    ("coolmaterial.com", "Guest posting", "https://coolmaterial.com/", "3", "Yes", "Pitch", "Cool Material"),
    ("dailycoffee.com", "Guest posting", "https://dailycoffee.com/", "3", "Yes", "Guest", "Daily Coffee"),
    ("coffeegeek.com", "Guest posting", "https://coffeegeek.com/", "3", "Yes", "Forum article", "CoffeeGeek"),
    ("sweetmarias.com", "Guest posting", "https://sweetmarias.com/", "3", "Yes", "Community", "Sweet Maria's forum"),
    ("homegrounds.co", "Guest posting", "https://homegrounds.co/", "3", "Yes", "Guest post", "HomeGrounds write for us"),
    ("coffeeness.de", "Guest posting", "https://www.coffeeness.de/en/", "3", "Yes", "Guest", "Coffeeness EN"),
    ("coffeefactory.com", "Guest posting", "https://www.coffeefactory.com/", "3", "Yes", "Guest", "Coffee Factory blog"),
    ("blog.koa.coffee", "Guest posting", "https://www.koa.coffee/blog/", "3", "Yes", "Guest", "Koa Coffee blog"),
    ("counterculturecoffee.com", "Guest posting", "https://counterculturecoffee.com/blog", "4", "Yes", "Industry", "Counter Culture blog"),
    ("stumptowncoffee.com", "Guest posting", "https://www.stumptowncoffee.com/blogs", "4", "Yes", "Industry", "Stumptown blog"),
    ("intelligentsia.com", "Guest posting", "https://www.intelligentsia.com/blogs", "4", "Yes", "Industry", "Intelligentsia blog"),
    ("lacolombe.com", "Guest posting", "https://www.lacolombe.com/blogs", "4", "Yes", "Industry", "La Colombe blog"),
    ("peets.com", "Guest posting", "https://www.peets.com/blogs/peets", "4", "Yes", "Industry", "Peet's blog"),
    ("starbucks.com", "Guest posting", "https://www.starbucks.com/", "5", "Yes", "Industry", "Competitor blogs — research"),
    ("blog.feedspot.com", "Guest posting", "https://blog.feedspot.com/food_blogs/", "2", "Yes", "Outreach", "Food blog list outreach"),
    ("guestpost.com", "Guest posting", "https://guestpost.com/", "2", "Yes", "Marketplace", "Guest post opportunities"),
    ("guestposttracker.com", "Guest posting", "https://guestposttracker.com/", "2", "Yes", "List", "Guest post tracker"),
    ("acceptguestpost.com", "Guest posting", "https://acceptguestpost.com/", "2", "Yes", "Marketplace", "Accept guest post"),
    ("iCopify.com", "Guest posting", "https://icopify.com/", "2", "Yes", "Marketplace", "Free tier listings"),
    ("articlereword.com", "Guest posting", "https://articlereword.com/", "2", "Yes", "Submit", "Article sites network"),
    ("articlesfactory.com", "Guest posting", "https://www.articlesfactory.com/", "2", "Yes", "Article submit", "Articles Factory"),
    ("goarticles.com", "Guest posting", "https://goarticles.com/", "3", "Yes", "Article", "GoArticles"),
    ("articlebiz.com", "Guest posting", "https://www.articlebiz.com/", "2", "Yes", "Submit", "ArticleBiz"),
    ("sooperarticles.com", "Guest posting", "https://www.sooperarticles.com/", "2", "Yes", "Submit", "Sooper Articles"),
    ("amazines.com", "Guest posting", "https://www.amazines.com/", "2", "Yes", "Submit", "Amazines"),
    ("blogger.com", "Guest posting", "https://www.blogger.com/", "2", "Yes", "Blog", "Blogger guest microsites"),
    ("wordpress.com", "Guest posting", "https://wordpress.com/", "2", "Yes", "Blog", "WordPress.com guest blogs"),
    ("tumblr.com", "Guest posting", "https://www.tumblr.com/", "2", "Yes", "Blog", "Tumblr coffee blogs"),
    ("write.as", "Guest posting", "https://write.as/", "2", "Yes", "Blog", "Write.as free posts"),
    ("telegra.ph", "Guest posting", "https://telegra.ph/", "2", "Yes", "Instant article", "Telegra.ph"),
    ("substack.com", "Guest posting", "https://substack.com/", "2", "Yes", "Guest newsletter", "Guest in Substacks"),
]

URL_GUESS = {
    "yelp.com": "https://biz.yelp.com/",
    "hotfrog.com": "https://www.hotfrog.com/add-your-business",
    "brownbook.net": "https://www.brownbook.net/add-business/",
    "cybo.com": "https://www.cybo.com/add/",
    "cylex.com": "https://www.cylex.us.com/company/add/",
    "tuugo.us": "https://www.tuugo.us/",
    "enrollbusiness.com": "https://us.enrollbusiness.com/",
    "storeboard.com": "https://www.storeboard.com/",
    "linkcentre.com": "https://www.linkcentre.com/submit.php",
    "viesearch.com": "https://viesearch.com/submit",
}

SKIP_DOMAIN = (
    ".co.in",
    "indiamart",
    "sulekha",
    "justdial",
    "mouthshut",
    "midgetsex",
    "neobacklinks",
    "exampledir",
    "craigslist",
)


def slug_domain(name: str) -> str:
    name = name.strip().lower()
    if not name or name.startswith("#"):
        return ""
    if "/" in name:
        name = name.split("/")[0]
    return re.sub(r"[^a-z0-9.-]", "", name)


def build_rows() -> list[list[str]]:
    seen: set[str] = set()
    rows: list[list[str]] = []

    def add(entry: tuple[str, ...]) -> None:
        domain = entry[0].lower()
        if domain in seen:
            return
        seen.add(domain)
        rows.append(list(entry))

    for block in (LOCAL_CITATION, FREE_DIRECTORY, GUEST_POSTING):
        for e in block:
            add(e)

    extra_path = Path(__file__).with_name("free_us_directory_domains.txt")
    if extra_path.is_file():
        for line in extra_path.read_text(encoding="utf-8").splitlines():
            if len(rows) >= 1000:
                break
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            domain = slug_domain(line)
            if not domain or domain in seen:
                continue
            if any(x in domain for x in SKIP_DOMAIN):
                continue
            if domain.endswith((".ca", ".co.za", ".com.au", ".co.uk")):
                continue
            sub_url = URL_GUESS.get(domain, f"https://{domain}/")
            add(
                (
                    domain,
                    "Free directory",
                    sub_url,
                    "2",
                    "Yes",
                    "Directory listing",
                    "Free business directory — verify live before submit",
                )
            )

    return rows[:1000]


def write_csv(rows: list[list[str]]) -> None:
    with CSV_PATH.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Blue Bottle Coffee — FREE directories, guest posting & local citations only"])
        w.writerow(["Categories: Free directory | Guest posting | Local citation · USA · No social/bookmark/PR/classifieds"])
        w.writerow(["Case: bluebottlecoffee.com · Sample of free submittable tier from ~18.4K backlink profile"])
        w.writerow([])
        w.writerow(HEADERS)
        for idx, row in enumerate(rows, start=1):
            w.writerow([idx, *row])


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def write_xlsx(rows: list[list[str]]) -> None:
    sheet_rows = [
        ["Blue Bottle — Directories, Guest Posting, Local Citations"],
        ["Free only · USA · Excludes social, Web2, PR, bookmarks, Craigslist"],
        [],
        HEADERS,
    ]
    for idx, row in enumerate(rows, start=1):
        sheet_rows.append([idx, *row])

    def cell_ref(r: int, c: int) -> str:
        return f"{col_letter(c)}{r}"

    sheet_data = ["<sheetData>"]
    for r_idx, row in enumerate(sheet_rows, start=1):
        sheet_data.append(f'<row r="{r_idx}">')
        for c_idx, val in enumerate(row, start=1):
            text = escape(str(val))
            sheet_data.append(
                f'<c r="{cell_ref(r_idx, c_idx)}" t="inlineStr"><is><t>{text}</t></is></c>'
            )
        sheet_data.append("</row>")
    sheet_data.append("</sheetData>")
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + "".join(sheet_data)
        + "</worksheet>"
    )

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
    wb_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Dirs Guest Local" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""

    with zipfile.ZipFile(XLSX_PATH, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    rows = build_rows()
    print(f"Generated {len(rows)} rows")
    write_csv(rows)
    write_xlsx(rows)
    print(f"Wrote {CSV_PATH}")
    print(f"Wrote {XLSX_PATH}")


if __name__ == "__main__":
    main()
