
// https://learn.microsoft.com/en-us/graph/api/resources/person?view=graph-rest-1.0
export interface IMicrosoftPeopleResponse {
    "value": IMicrosoftPeopleResultSet[];
}

export interface IMicrosoftPeopleResultSet {
        "id": string;
        "displayName": string;
        "givenName": string;
        "surname": string;
        "birthday": string;
        "personNotes": string;
        "isFavorite": boolean;
        "title": string;
        "companyName": string;
        "yomiCompany": string;
        "department": string;
        "officeLocation": string;
        "profession": string;
        "mailboxType": string;
        "personType": string;
        "userPrincipalName": string;
        "emailAddresses": IPeopleEmailAddress[];
        "phones": IPeoplePhone[];
        "postalAddresses": IPeoplePostalAddress[];
        "websites": IPeopleWebsite[];
        "sources": IPeopleSource[];
}

export interface IPeopleEmailAddress {
    "address": string;
    "rank": number;
}

export interface IPeoplePhone {
    "number": string;
    "type": string;
}

export interface IPeoplePostalAddress {
    "address": IPeoplePostalAddressesAddress,
    "coordinates": { "@odata.type": "microsoft.graph.outlookGeoCoordinates" },
    "displayName": string,
    "locationEmailAddress": string,
    "locationType": string,
    "locationUri": string,
    "uniqueId": string,
    "uniqueIdType": string
}

export interface IPeopleWebsite {
    "type": string;
    "address": string;
    "displayName": string;
}

export interface IPeopleSource {
    "type": string;
}

export interface IPeoplePostalAddressesAddress {
    "city": string;
    "countryOrRegion": string;
    "postalCode": string;
    "postOfficeBox": string;
    "state": string;
    "street": string;
    "type": string;
}

export interface IPeoplePostalAddressesCoordinates {
    "accuracy": number;
    "altitude": number;
    "altitudeAccuracy": number;
    "latitude": number;
    "longitude": number;
}