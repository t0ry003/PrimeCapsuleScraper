import os
import requests
import xlwt
from xlwt import Workbook


# set the style for the Excel sheet
def set_style(font_name, bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    font.outline = True
    font.name = font_name
    font.bold = bold
    font.color_index = 4
    style.font = font
    return style


# check if file exists then delete it
def check_file(file_name):
    try:
        if os.path.exists(file_name):
            os.remove(file_name)
            print(f"Old \"{file_name}\" deleted")
    except OSError:
        print(f"Error: can't delete {file_name}")


url = "https://gaming.amazon.com/graphql"

querystring = {"nonce": "7c883189-cc2d-439b-903f-b6d1d8743eb6"}

payload = {
    "operationName": "OfferDetail_Journey",
    "variables": {
        "journeyId": "5f82d116-ccd0-4595-94b8-77052153fe57",
        "redirectUrl": "https://gaming.amazon.com/loot/leagueoflegends",
        "journeyShortName": "leagueoflegends",
        "stringDebug": False
    },
    "query": """fragment OfferDetail_Journey_Media on Media {
        alt
        defaultMedia {
            src1x
            src2x
            __typename
        }
        desktop {
            src1x
            src2x
            __typename
        }
        tablet {
            src1x
            src2x
            __typename
        }
        __typename
        }

        fragment JourneyHeroAsset on MediaAsset {
        src1x
        src2x
        type
        __typename
        }

        fragment OfferDetail_Journey_Offer_Pixel on Pixel {
        type
        pixel
        __typename
        }

        fragment OfferDetail_Journey_Offer on JourneyOffer {
        catalogId
        id
        startTime
        endTime
        grantsCode
        isFGWP
        self {
            claimStatus
            orderInformation {
            ...OfferDetail_Journey_Offer_OrderInformation
            __typename
            }
            eligibility {
            isClaimed
            canClaim
            isPrimeGaming
            missingRequiredAccountLink
            gameAccountDisplayName
            offerStartTime
            offerEndTime
            offerState
            inRestrictedMarketplace
            maxOrdersExceeded
            conflictingClaimAccount {
                ...ConflictingClaimAccount
                __typename
            }
            conflictingThirdPartyAccounts {
                accountType
                name
                __typename
            }
            __typename
            }
            __typename
        }
        assets {
            additionalMedia {
            ...OfferDetail_Journey_Media
            __typename
            }
            card {
            ...OfferDetail_Journey_Media
            __typename
            }
            items
            subtitle
            externalClaimLink
            title
            header
            pixels {
            ...OfferDetail_Journey_Offer_Pixel
            __typename
            }
            __typename
        }
        __typename
        }

        fragment JourneyContextHeroAssets on Media {
        defaultMedia {
            ...JourneyHeroAsset
            __typename
        }
        tablet {
            ...JourneyHeroAsset
            __typename
        }
        desktop {
            ...JourneyHeroAsset
            __typename
        }
        alt
        __typename
        }

        fragment ActionsContext_JourneyAssets on JourneyAssets {
        claimVisualInstructions {
            ...OfferDetail_Journey_Media
            __typename
        }
        hero {
            ...JourneyContextHeroAssets
            __typename
        }
        title
        subtitle
        description
        platformAvailability
        publisherName
        thirdPartyAccountManagementUrl
        claimInstructions
        purchaseGameText
        __typename
        }

        fragment OfferDetail_Journey_Offer_OrderInformation on OfferOrderInformation {
        orderDate
        orderState
        claimCode
        __typename
        }

        fragment OfferDetail_Journey_Game on GameV2 {
        id
        assets {
            title
            __typename
        }
        gameSelfConnection {
            isSubscribedToNotifications
            __typename
        }
        __typename
        }

        query OfferDetail_Journey($journeyId: String!, $redirectUrl: String!, $journeyShortName: String!, 
        $dateOverride: Time, $stringDebug: Boolean) { journey(journeyId: $journeyId, dateOverride: $dateOverride, 
        stringDebug: $stringDebug) { id offers { ...OfferDetail_Journey_Offer __typename } assets { 
        ...ActionsContext_JourneyAssets __typename } accountLinkConfig(redirectUrl: $redirectUrl, journeyShortName: 
        $journeyShortName, stringDebug: $stringDebug) { accountType linkedAccountConfirmation linkingInstructions 
        linkingUrl __typename } game { ...OfferDetail_Journey_Game __typename } __typename } } 

        fragment ConflictingClaimAccount on ConflictingClaimUser {
        fullName
        obfuscatedEmail
        __typename
        }
""",
    "extensions": {}
}
headers = {
    "cookie": "session-id=141-5621786-7681141; session-id-time=2082787201l; "
              "session-token=CqPxDdSyXqJQPpKdnXo69XZbMfrbR5R9kHmWgu6nyp9ePyI9EgKGBA9zKPuYH1BIpThOB4ENNqmSWlJ"
              "%2Bb9QUotQCzV2J0sa5U%2Bptf2Fe0NUFYyB1SJOS%2Btg55H3NimTIQXs8OGjbhPLP3w8s%2FulFX"
              "%2BAnBcmTDGlcf6MvZHi3ye19lqx0qHEy43rKsPeOsTaQpsIfnUWViMWC5XFlaMsQJQ",
    "Host": "gaming.amazon.com",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:102.0) Gecko/20100101 Firefox/102.0",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.5",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": "https://gaming.amazon.com/loot/leagueoflegends",
    "Origin": "https://gaming.amazon.com",
    "Connection": "keep-alive",
    "Cookie": "session-id=141-5621786-7681141; session-id-time=2082787201l; unique_id=0ee709a8-d203-41; "
              "ubid-main=135-8392164-9500608; twitch-prime-language=en-US; "
              "session-token=CqPxDdSyXqJQPpKdnXo69XZbMfrbR5R9kHmWgu6nyp9ePyI9EgKGBA9zKPuYH1BIpThOB4ENNqmSWlJ"
              "+b9QUotQCzV2J0sa5U+ptf2Fe0NUFYyB1SJOS+tg55H3NimTIQXs8OGjbhPLP3w8s/ulFX"
              "+AnBcmTDGlcf6MvZHi3ye19lqx0qHEy43rKsPeOsTaQpsIfnUWViMWC5XFlaMsQJQ",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "no-cors",
    "Sec-Fetch-Site": "same-origin",
    "Content-Type": "application/json",
    "Client-Id": "CarboniteApp",
    "csrf-token": "gknwvIqxqFzV1kM8gxT7veEpq9C1/7OGSNXQstgAAAACAAAAAGLCyiFyYXcAAAAA+8jokd9rqj+wHxPcX6iU",
    "prime-gaming-language": "en-US",
    "Pragma": "no-cache",
    "Cache-Control": "no-cache"
}

response = requests.request("POST", url, json=payload, headers=headers, params=querystring)

wb = Workbook()

name = "League of Legends Capsules"

sheet = wb.add_sheet("League of Legends Capsules")

sheet.write(0, 1, "Start Time", set_style('Arial', bold=True))
sheet.write(0, 2, "End Time", set_style('Arial', bold=True))

index = 1
for offer in response.json()['data']['journey']['offers']:
    sheet.write(index, 0, index)
    sheet.write(index, 1, offer['startTime'])
    sheet.write(index, 2, offer['endTime'])
    index += 1

check_file(f"{name}.xls")
wb.save(f"{name}.xls")

print(f"New \"{name}.xls\" saved successfully")
