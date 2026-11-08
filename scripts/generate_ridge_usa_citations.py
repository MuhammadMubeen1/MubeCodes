#!/usr/bin/env python3
"""USA-only free citation sites for local businesses.

Est. DA / PA / spam score are not a live Moz pull. BrightLocal figures are
used where that list names the site. Everything else is a conservative
estimate — confirm in Moz before submission.
Spam score is Moz-style: 1–100, lower is cleaner.
"""
from __future__ import annotations

import csv
import socket
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(__file__).resolve().parents[1] / "portfolio" / "reports"
CSV_PATH = OUT / "ridge-wallet-usa-free-citations.csv"
XLSX_PATH = OUT / "ridge-wallet-usa-free-citations.xlsx"

HEADERS = [
    "Referring domain",
    "Submission URL",
]

# domain, type, da, pa, spam
# USA businesses only. Foreign directories are not in this list.
SITES: list[tuple[str, str, int, int, int]] = [
    ("business.google.com", "Maps", 100, 86, 1),
    ("businessconnect.apple.com", "Maps", 99, 78, 1),
    ("linkedin.com", "Business profile", 99, 80, 1),
    ("merchants.google.com", "Shopping", 98, 70, 1),
    ("facebook.com", "Business profile", 94, 74, 1),
    ("webmd.com", "Health", 94, 60, 1),
    ("yelp.com", "Reviews", 93, 70, 1),
    ("bingplaces.com", "Maps", 93, 68, 1),
    ("instagram.com", "Business profile", 93, 66, 1),
    ("tripadvisor.com", "Food and travel", 93, 60, 1),
    ("indeed.com", "Company profile", 92, 62, 1),
    ("glassdoor.com", "Company profile", 92, 60, 1),
    ("waze.com", "Maps", 92, 55, 1),
    ("foursquare.com", "Maps", 91, 64, 1),
    ("bbb.org", "Reviews", 91, 62, 1),
    ("crunchbase.com", "Company profile", 91, 62, 1),
    ("justia.com", "Legal", 90, 58, 1),
    ("patch.com", "Local community", 90, 55, 2),
    ("producthunt.com", "Company profile", 90, 58, 2),
    ("realtor.com", "Real estate", 90, 58, 1),
    ("zillow.com", "Real estate", 89, 58, 1),
    ("findlaw.com", "Legal", 88, 56, 2),
    ("mapquest.com", "Maps", 88, 58, 2),
    ("trulia.com", "Real estate", 86, 54, 1),
    ("yellowpages.com", "General directory", 86, 58, 2),
    ("houzz.com", "Home services", 86, 55, 1),
    ("redfin.com", "Real estate", 85, 52, 1),
    ("consumeraffairs.com", "Reviews", 84, 56, 2),
    ("theknot.com", "Weddings", 84, 54, 2),
    ("wellfound.com", "Company profile", 84, 54, 2),
    ("dnb.com", "Company profile", 82, 55, 2),
    ("whitepages.com", "General directory", 82, 54, 3),
    ("nextdoor.com", "Local community", 83, 54, 1),
    ("nolo.com", "Legal", 80, 50, 2),
    ("manta.com", "General directory", 80, 52, 6),
    ("ziprecruiter.com", "Company profile", 78, 50, 2),
    ("homeadvisor.com", "Home services", 78, 50, 2),
    ("untappd.com", "Food and drink", 78, 48, 2),
    ("martindale.com", "Legal", 78, 50, 2),
    ("care.com", "Home and care", 77, 48, 2),
    ("monster.com", "Company profile", 76, 48, 3),
    ("merchantcircle.com", "General directory", 76, 50, 8),
    ("angi.com", "Home services", 76, 50, 2),
    ("archinect.com", "Architecture", 76, 48, 3),
    ("porch.com", "Home services", 76, 48, 3),
    ("ratebeer.com", "Food and drink", 76, 46, 4),
    ("thumbtack.com", "Home services", 76, 48, 2),
    ("zoominfo.com", "Company profile", 75, 50, 2),
    ("apsense.com", "General directory", 75, 46, 28),
    ("lawyers.com", "Legal", 74, 48, 3),
    ("thomasnet.com", "Industry", 74, 50, 2),
    ("builtin.com", "Company profile", 74, 48, 2),
    ("sitejabber.com", "Reviews", 74, 48, 4),
    ("avvo.com", "Legal", 74, 48, 2),
    ("careerbuilder.com", "Company profile", 73, 48, 3),
    ("sharecare.com", "Health", 72, 46, 2),
    ("taskrabbit.com", "Home services", 72, 46, 2),
    ("uslegal.com", "Legal", 72, 46, 5),
    ("homes.com", "Real estate", 72, 46, 2),
    ("weddingwire.com", "Weddings", 72, 48, 2),
    ("superpages.com", "General directory", 71, 48, 3),
    ("dice.com", "Company profile", 71, 46, 3),
    ("happycow.com", "Food and drink", 70, 44, 3),
    ("citysearch.com", "General directory", 70, 46, 5),
    ("simplyhired.com", "Company profile", 70, 46, 3),
    ("batchgeo.com", "Maps", 70, 40, 4),
    ("realself.com", "Health", 70, 44, 3),
    ("bestplumbers.com", "Home services", 70, 42, 16),
    ("dexknows.com", "General directory", 69, 46, 4),
    ("healthgrades.com", "Health", 69, 46, 2),
    ("alignable.com", "Local community", 68, 48, 3),
    ("411.com", "General directory", 68, 46, 5),
    ("craft.co", "Company profile", 68, 46, 3),
    ("vitals.com", "Health", 68, 44, 2),
    ("zeemaps.com", "Maps", 67, 45, 8),
    ("faire.com", "Industry", 67, 46, 3),
    ("owler.com", "Company profile", 66, 45, 4),
    ("hg.org", "Legal", 66, 42, 4),
    ("chamberofcommerce.com", "General directory", 65, 46, 7),
    ("elocal.com", "General directory", 64, 45, 11),
    ("storeboard.com", "General directory", 64, 45, 22),
    ("comparably.com", "Company profile", 64, 45, 3),
    ("anywho.com", "General directory", 63, 45, 6),
    ("resellerratings.com", "Reviews", 62, 45, 6),
    ("buildzoom.com", "Home services", 62, 40, 4),
    ("allmenus.com", "Food and drink", 62, 40, 6),
    ("lawinfo.com", "Legal", 62, 40, 6),
    ("wellness.com", "Health", 62, 40, 4),
    ("yellowbook.com", "General directory", 61, 45, 6),
    ("handy.com", "Home services", 60, 40, 4),
    ("local.com", "General directory", 60, 45, 12),
    ("spoke.com", "General directory", 59, 40, 9),
    ("networx.com", "Home services", 58, 38, 6),
    ("hubbiz.com", "General directory", 55, 36, 10),
    ("trepup.com", "General directory", 57, 36, 16),
    ("americantowns.com", "General directory", 56, 38, 8),
    ("ezlocal.com", "General directory", 55, 38, 10),
    ("callupcontact.com", "General directory", 55, 36, 24),
    ("insiderpages.com", "General directory", 55, 36, 10),
    ("twidloo.com", "General directory", 55, 34, 20),
    ("yplocal.com", "General directory", 55, 36, 12),
    ("modernize.com", "Home services", 55, 36, 8),
    ("ebusinesspages.com", "General directory", 54, 36, 16),
    ("us.enrollbusiness.com", "General directory", 54, 34, 18),
    ("showmelocal.com", "General directory", 53, 36, 14),
    ("judysbook.com", "General directory", 52, 36, 8),
    ("bubblelife.com", "Local community", 50, 34, 12),
    ("salespider.com", "General directory", 50, 32, 20),
    ("citysquares.com", "General directory", 48, 34, 12),
    ("cityof.com", "General directory", 48, 33, 15),
    ("parkbench.com", "General directory", 48, 32, 14),
    ("agreatertown.com", "General directory", 48, 32, 16),
    ("yellowbot.com", "General directory", 48, 32, 14),
    ("homeyou.com", "Home services", 48, 32, 10),
    ("menupix.com", "Food and drink", 48, 32, 10),
    ("improvenet.com", "Home services", 48, 32, 12),
    ("2findlocal.com", "General directory", 47, 32, 18),
    ("cityfos.com", "General directory", 46, 30, 16),
    ("place123.net", "General directory", 46, 30, 28),
    ("craftjack.com", "Home services", 45, 30, 12),
    ("kudzu.com", "General directory", 45, 30, 12),
    ("find-us-here.com", "General directory", 44, 30, 22),
    ("uscity.net", "General directory", 44, 30, 18),
    ("ibegin.com", "General directory", 42, 28, 18),
    ("localstack.com", "General directory", 42, 28, 16),
    ("n2local.com", "General directory", 40, 28, 20),
    ("getfave.com", "General directory", 40, 26, 15),
    ("yasabe.com", "General directory", 38, 26, 14),
    ("bizhwy.com", "General directory", 37, 24, 22),
    ("golocal247.com", "General directory", 36, 24, 18),
    ("localpages.com", "General directory", 35, 24, 20),
    ("expressbusinessdirectory.com", "General directory", 32, 20, 35),
    ("zipleaf.us", "General directory", 30, 20, 30),
    ("freelistingusa.com", "General directory", 28, 18, 42),
    ("youtube.com", "Business profile", 99, 70, 1),
    ("x.com", "Business profile", 96, 64, 2),
    ("pinterest.com", "Business profile", 94, 60, 1),
    ("vimeo.com", "Business profile", 90, 55, 2),
    ("health.usnews.com", "Health", 90, 58, 1),
    ("groupon.com", "Local deals", 88, 54, 2),
    ("carfax.com", "Auto", 86, 54, 1),
    ("cars.com", "Auto", 84, 52, 1),
    ("fodors.com", "Food and travel", 82, 52, 1),
    ("kbb.com", "Auto", 82, 52, 1),
    ("edmunds.com", "Auto", 80, 50, 1),
    ("greatschools.org", "Education", 80, 52, 1),
    ("autotrader.com", "Auto", 78, 50, 2),
    ("niche.com", "Education", 78, 50, 2),
    ("hotpads.com", "Real estate", 76, 48, 2),
    ("rover.com", "Pets", 76, 50, 2),
    ("frommers.com", "Food and travel", 74, 48, 2),
    ("angieslist.com", "Home services", 74, 48, 2),
    ("zocdoc.com", "Health", 74, 48, 2),
    ("superlawyers.com", "Legal", 72, 48, 2),
    ("doximity.com", "Health", 72, 46, 2),
    ("snagajob.com", "Company profile", 70, 46, 3),
    ("aplaceformom.com", "Health", 70, 46, 2),
    ("wagwalking.com", "Pets", 70, 46, 2),
    ("globalspec.com", "Industry", 70, 46, 3),
    ("classpass.com", "Fitness", 70, 46, 2),
    ("caring.com", "Health", 68, 44, 2),
    ("mindbodyonline.com", "Fitness", 68, 44, 3),
    ("zagat.com", "Food and drink", 68, 44, 2),
    ("rent.com", "Real estate", 68, 44, 2),
    ("styleseat.com", "Beauty", 66, 42, 2),
    ("doctor.com", "Health", 66, 42, 2),
    ("castleconnolly.com", "Health", 64, 42, 2),
    ("allaboutvision.com", "Health", 64, 42, 2),
    ("agingcare.com", "Health", 64, 42, 3),
    ("apartmentlist.com", "Real estate", 64, 42, 2),
    ("bringfido.com", "Pets", 62, 40, 3),
    ("goodtherapy.org", "Health", 62, 40, 3),
    ("repairpal.com", "Auto", 62, 40, 3),
    ("loc8nearme.com", "General directory", 62, 42, 6),
    ("corporationwiki.com", "Company profile", 62, 40, 4),
    ("vagaro.com", "Beauty", 60, 40, 3),
    ("menupages.com", "Food and drink", 60, 40, 6),
    ("homesnap.com", "Real estate", 60, 40, 3),
    ("industrynet.com", "Industry", 60, 40, 4),
    ("company.com", "General directory", 60, 40, 8),
    ("scribblemaps.com", "Maps", 60, 40, 6),
    ("switchboard.com", "General directory", 60, 40, 6),
    ("hub.biz", "General directory", 58, 38, 8),
    ("yourmechanic.com", "Auto", 58, 38, 3),
    ("bizapedia.com", "Company profile", 58, 38, 5),
    ("senioradvisor.com", "Health", 58, 38, 3),
    ("myhuckleberry.com", "General directory", 58, 38, 8),
    ("communitywalk.com", "Maps", 58, 38, 8),
    ("dentalplans.com", "Health", 56, 36, 4),
    ("hvac.com", "Home services", 56, 36, 6),
    ("betalist.com", "Company profile", 58, 36, 4),
    ("cortera.com", "Company profile", 55, 36, 8),
    ("yellowpagecity.com", "General directory", 55, 36, 12),
    ("parishesonline.com", "Local community", 55, 36, 6),
    ("guildquality.com", "Home services", 54, 36, 4),
    ("dentistry.com", "Health", 52, 34, 6),
    ("finduslocal.com", "General directory", 52, 34, 12),
    ("40billion.com", "General directory", 52, 36, 18),
    ("shopperapproved.com", "Reviews", 52, 34, 6),
    ("landscapingnetwork.com", "Home services", 52, 34, 6),
    ("townplanner.com", "Local community", 50, 34, 10),
    ("claimspages.com", "Insurance", 50, 34, 10),
    ("infofree.com", "Company profile", 50, 32, 16),
    ("b2byellowpages.com", "General directory", 48, 32, 16),
    ("igotbiz.com", "General directory", 48, 32, 18),
    ("whatsyourhours.com", "General directory", 48, 32, 12),
    ("areaconnect.com", "General directory", 48, 32, 12),
    ("localdatabase.com", "General directory", 48, 32, 14),
    ("solvhealth.com", "Health", 48, 32, 6),
    ("weddingspot.com", "Weddings", 48, 32, 6),
    ("carwise.com", "Auto", 48, 32, 8),
    ("gayot.com", "Food and drink", 48, 32, 8),
    ("welocalpeople.com", "General directory", 46, 30, 14),
    ("schedulicity.com", "Beauty", 46, 30, 6),
    ("herecomestheguide.com", "Weddings", 46, 30, 6),
    ("gomylocal.com", "General directory", 44, 28, 16),
    ("peeplocal.com", "General directory", 44, 28, 16),
    ("businesslistingplus.com", "General directory", 44, 28, 20),
    ("massagebook.com", "Beauty", 44, 28, 6),
    ("mechanicadvisor.com", "Auto", 44, 28, 8),
    ("jasminedirectory.com", "General directory", 42, 28, 20),
    ("hrbusinesslive.com", "General directory", 42, 28, 18),
    ("localmint.com", "General directory", 42, 28, 16),
    ("nextbizthing.com", "General directory", 42, 28, 20),
    ("genbook.com", "Beauty", 42, 28, 8),
    ("lawguru.com", "Legal", 42, 28, 10),
    ("legalmatch.com", "Legal", 46, 30, 8),
    ("webforcompany.com", "General directory", 40, 26, 22),
    ("expertondeck.com", "Home services", 40, 28, 12),
    ("directorytogoto.com", "General directory", 40, 26, 22),
    ("eliteservicesnetwork.com", "General directory", 40, 26, 22),
    ("bizbangboom.com", "General directory", 40, 26, 24),
    ("ilocatelocal.com", "General directory", 40, 26, 18),
    ("homeblue.com", "Home services", 40, 26, 14),
    ("proreferral.com", "Home services", 42, 28, 10),
    ("citylocalpro.com", "General directory", 36, 24, 18),
    ("cityinsider.com", "General directory", 36, 24, 16),
    ("gymlocator.com", "Fitness", 36, 24, 12),
    ("localpacket.com", "General directory", 34, 22, 20),
    ("ypages.us", "General directory", 32, 20, 18),
    ("gatordirectory.com", "General directory", 30, 20, 24),
    ("pizza.com", "Food and drink", 34, 22, 20),
    ("getlistedusa.com", "General directory", 28, 18, 20),
    ("usalync.com", "General directory", 24, 16, 22),
    ("schooldigger.com", "Education", 58, 38, 4),
    ("publicschoolreview.com", "Education", 62, 40, 3),
    ("privateschoolreview.com", "Education", 55, 36, 4),
    ("macraesbluebook.com", "Industry", 52, 34, 8),
    ("iqsdirectory.com", "Industry", 48, 32, 8),
    ("coffeeshopdirectory.com", "Food and drink", 28, 18, 24),
    ("wineryhunt.com", "Food and drink", 26, 16, 22),
    ("plumbersnearyou.com", "Home services", 24, 16, 28),
    ("personalinjurylawyers.com", "Legal", 30, 20, 22),
    ("eyecare.org", "Health", 40, 26, 8),
    ("chiropractordirectory.com", "Health", 22, 14, 26),
    ("offerup.com", "Local marketplace", 72, 46, 3),
    ("seamless.com", "Food and drink", 70, 44, 3),
    ("servicemagic.com", "Home services", 55, 36, 6),
    ("opentbay.com", "Auto", 42, 28, 8),
    ("md.com", "Health", 48, 32, 6),
    ("amino.com", "Health", 55, 36, 4),
    ("8coupons.com", "Local deals", 48, 32, 10),
    ("infospace.com", "General directory", 52, 34, 10),
    ("jayde.com", "General directory", 42, 28, 16),
    ("exactseek.com", "General directory", 38, 26, 18),
    ("somuch.com", "General directory", 36, 24, 20),
    ("bunity.com", "General directory", 36, 24, 22),
    ("addonbiz.com", "General directory", 28, 20, 28),
    ("a-zbusinessfinder.com", "General directory", 26, 20, 30),
    ("usalistingdirectory.com", "General directory", 26, 20, 32),
    ("allstatesusadirectory.com", "General directory", 24, 20, 34),
    ("usawebsitesdirectory.com", "General directory", 24, 20, 34),
    ("411.info", "General directory", 40, 26, 12),
    ("mojopages.com", "General directory", 34, 22, 16),
    ("cityvoter.com", "General directory", 32, 22, 18),
    ("askmap.net", "Maps", 30, 20, 22),
    ("companylistingnyc.com", "General directory", 36, 24, 16),
    ("pr.business", "General directory", 40, 26, 18),
    ("visitcalifornia.com", "Local citation", 74, 50, 2),
    ("nyctourism.com", "Local citation", 72, 48, 2),
    ("choosechicago.com", "Local citation", 70, 46, 2),
    ("washington.org", "Local citation", 70, 46, 2),
    ("virginia.org", "Local citation", 68, 44, 2),
    ("sftravel.com", "Local citation", 68, 46, 2),
    ("visitphilly.com", "Local citation", 68, 44, 2),
    ("visitseattle.org", "Local citation", 66, 44, 2),
    ("miamiandbeaches.com", "Local citation", 66, 44, 2),
    ("traveltexas.com", "Local citation", 66, 44, 2),
    ("visitdallas.com", "Local citation", 64, 42, 2),
    ("denver.org", "Local citation", 64, 42, 2),
    ("exploregeorgia.org", "Local citation", 64, 42, 2),
    ("bostonusa.com", "Local citation", 63, 42, 2),
    ("austintexas.org", "Local citation", 62, 40, 2),
    ("locable.com", "Local citation", 50, 34, 8),
    ("menuism.com", "Food and drink", 48, 32, 8),
    ("localsearch.com", "General directory", 42, 28, 12),
    ("chamber.com", "General directory", 38, 26, 12),
    ("where2go.com", "Local citation", 36, 24, 12),
    ("guidetogwinnett.com", "Local citation", 36, 24, 8),
    ("findlocal.com", "General directory", 34, 22, 14),
    ("neighbor.com", "Local citation", 34, 22, 12),
    ("findmecoffee.com", "Food and drink", 32, 22, 10),
    ("texomaland.com", "Local citation", 32, 22, 8),
    ("alabamaindex.com", "General directory", 30, 20, 14),
    ("idahoindex.com", "General directory", 30, 20, 14),
]

URLS = {
    "business.google.com": "https://business.google.com/",
    "businessconnect.apple.com": "https://businessconnect.apple.com/",
    "bingplaces.com": "https://www.bingplaces.com/",
    "merchants.google.com": "https://merchants.google.com/",
    "yelp.com": "https://business.yelp.com/",
    "bbb.org": "https://www.bbb.org/get-listed",
    "facebook.com": "https://www.facebook.com/pages/create/",
    "instagram.com": "https://business.instagram.com/",
    "linkedin.com": "https://www.linkedin.com/company/setup/new/",
    "nextdoor.com": "https://business.nextdoor.com/",
    "manta.com": "https://www.manta.com/add-your-business",
    "merchantcircle.com": "https://www.merchantcircle.com/signup",
    "mapquest.com": "https://listings.mapquest.com/",
    "foursquare.com": "https://foursquare.com/",
    "angi.com": "https://www.angi.com/",
    "thumbtack.com": "https://www.thumbtack.com/pro",
    "houzz.com": "https://www.houzz.com/pro",
    "homeadvisor.com": "https://www.homeadvisor.com/",
    "chamberofcommerce.com": "https://www.chamberofcommerce.com/add-your-business",
    "indeed.com": "https://www.indeed.com/hire/company-pages",
    "glassdoor.com": "https://www.glassdoor.com/employers/",
    "dnb.com": "https://www.dnb.com/duns-number.html",
    "us.enrollbusiness.com": "https://us.enrollbusiness.com/",
    "cityfos.com": "https://www.cityfos.com/",
    "find-us-here.com": "https://www.find-us-here.com/",
    "waze.com": "https://www.waze.com/",
}

SOCIAL = {
    "facebook.com",
    "instagram.com",
    "linkedin.com",
    "youtube.com",
    "x.com",
    "twitter.com",
    "pinterest.com",
    "vimeo.com",
    "tiktok.com",
    "snapchat.com",
    "threads.net",
    "reddit.com",
    "tumblr.com",
    "nextdoor.com",
    "alignable.com",
}

FOREIGN_EXACT = {
    "hotfrog.com",
    "hotfrog.us",
    "brownbook.net",
    "cylex.com",
    "cylex.us.com",
    "tupalo.com",
    "cybo.com",
    "tuugo.us",
    "tuugo.com",
    "kompass.com",
    "provenexpert.com",
    "infobel.com",
    "trustpilot.com",
    "tomtom.com",
    "openstreetmap.org",
    "justlanded.com",
    "lacartes.com",
    "opendi.us",
    "zomato.com",
    "n49.com",
    "bark.com",
    "yell.com",
    "yellowpages.ca",
    "here.com",
}


def spam_band(score: int) -> str:
    if score <= 10:
        return "Low"
    if score <= 30:
        return "Medium"
    return "High"


def note_for(kind: str, spam: int) -> str:
    base = "Free USA listing. Use one business name, address, and phone."
    niche = {
        "Health",
        "Legal",
        "Real estate",
        "Home services",
        "Home and care",
        "Food and travel",
        "Food and drink",
        "Weddings",
        "Architecture",
        "Industry",
        "Shopping",
    }
    if kind in niche:
        base = f"{kind} directory. List the business only if this is its category. " + base
    if spam >= 31:
        base = "Higher spam risk. Use after Google, Apple, Bing, Yelp, and BBB. " + base
    elif spam >= 11:
        base = "Secondary USA directory. " + base
    return base


def resolves(domain: str) -> bool:
    hosts = [domain] if domain.startswith("www.") else [domain, f"www.{domain}"]
    for host in hosts:
        try:
            socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
            return True
        except socket.gaierror:
            continue
    return False


def build_rows() -> tuple[list[list[str]], list[str]]:
    dropped: list[str] = []
    candidates: list[tuple[str, str, int, int, int]] = []
    seen: set[str] = set()
    for domain, kind, da, pa, spam in SITES:
        domain = domain.lower().strip()
        if domain in seen:
            continue
        seen.add(domain)
        if domain in SOCIAL:
            dropped.append(f"{domain} social media")
            continue
        if domain in FOREIGN_EXACT or domain.endswith((".uk", ".ca", ".au", ".in", ".de", ".fr", ".pk", ".nz")):
            dropped.append(f"{domain} non-USA")
            continue
        if da < 20 or pa < 20:
            dropped.append(f"{domain} below DA 20 or PA 20")
            continue
        candidates.append((domain, kind, da, pa, spam))

    status: dict[str, bool] = {}
    with ThreadPoolExecutor(max_workers=24) as pool:
        futures = {pool.submit(resolves, row[0]): row[0] for row in candidates}
        for fut in as_completed(futures):
            domain = futures[fut]
            try:
                status[domain] = bool(fut.result())
            except Exception:
                status[domain] = False

    kept: list[list[str]] = []
    for domain, kind, da, pa, spam in candidates:
        if not status.get(domain):
            dropped.append(f"{domain} did not resolve")
            continue
        url = URLS.get(domain, f"https://{domain}/")
        kept.append(
            [
                domain,
                kind,
                url,
                "Free",
                str(da),
                str(pa),
                str(spam),
                spam_band(spam),
                "Yes",
                note_for(kind, spam),
            ]
        )
    kept.sort(key=lambda r: (int(r[4]), int(r[5]), r[0]))
    return kept, dropped


def export_rows(rows: list[list[str]]) -> list[list[str]]:
    return [[row[0], row[2]] for row in rows]


def write_csv(rows: list[list[str]]) -> None:
    with CSV_PATH.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(HEADERS)
        w.writerows(export_rows(rows))


def write_xlsx(rows: list[list[str]]) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        write_xlsx_fallback(rows)
        return

    exported = export_rows(rows)
    wb = Workbook()
    ws = wb.active
    ws.title = "USA Citations"
    ws.append(HEADERS)
    for row in exported:
        ws.append(row)

    head_font = Font(name="Calibri", bold=True, color="FFFFFF")
    fill_head = PatternFill("solid", fgColor="24556B")
    thin = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE"),
    )
    last_col = len(HEADERS)
    for col in range(1, last_col + 1):
        cell = ws.cell(1, col)
        cell.font = head_font
        cell.fill = fill_head
        cell.alignment = Alignment(vertical="center")
    for r in range(2, 2 + len(exported)):
        for c in range(1, last_col + 1):
            cell = ws.cell(r, c)
            cell.alignment = Alignment(vertical="center")
            cell.border = thin
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 52
    ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{1 + len(exported)}"
    ws.freeze_panes = "A2"
    wb.save(XLSX_PATH)


def write_xlsx_fallback(rows: list[list[str]]) -> None:
    sheet_rows = [HEADERS, *export_rows(rows)]

    def col_letter(n: int) -> str:
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    parts = ["<sheetData>"]
    for r_idx, row in enumerate(sheet_rows, start=1):
        parts.append(f'<row r="{r_idx}">')
        for c_idx, val in enumerate(row, start=1):
            text = escape(str(val))
            ref = f"{col_letter(c_idx)}{r_idx}"
            parts.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
        parts.append("</row>")
    parts.append("</sheetData>")
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + "".join(parts)
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
<sheets><sheet name="USA Citations" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""
    with zipfile.ZipFile(XLSX_PATH, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    rows, dropped = build_rows()
    write_csv(rows)
    write_xlsx(rows)
    print(f"Kept {len(rows)}")
    for line in dropped:
        print("DROP", line)
    print(CSV_PATH)
    print(XLSX_PATH)


if __name__ == "__main__":
    main()
