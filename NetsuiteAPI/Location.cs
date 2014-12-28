using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetsuiteAPI.com.netsuite.webservices;
using System.Net;


namespace NetsuiteAPI
{
    public class Location
    {
        public RecordRef locationRecord { get; set; }
        public Location()
        {
        }

        public Location(string locationName)
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            LocationSearch locSearch = new LocationSearch();
            SearchStringField locName = new SearchStringField();
            locName.@operator = SearchStringFieldOperator.@is;
            locName.operatorSpecified = true;
            locName.searchValue = locationName;
            LocationSearchBasic locSearchBasic = new LocationSearchBasic();
            locSearchBasic.name = locName;
            locSearch.basic = locSearchBasic;

            SearchResult locationResult = service.search(locSearch);

            if (locationResult.status.isSuccess != true) throw new Exception("Cannot find Item " + locationName + " " + locationResult.status.statusDetail[0].message);
            if (locationResult.recordList.Count() != 1) throw new Exception("More than one item found for item " + locationName);

            this.locationRecord = new RecordRef();
            this.locationRecord.type = RecordType.location;
            this.locationRecord.typeSpecified = true;
            this.locationRecord.internalId = ((com.netsuite.webservices.Location)locationResult.recordList[0]).internalId;
        }

    }
}

