using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp2
{
    public class BusinessInfo
    {
        public int? serialNumber {  get; set; }
        public string SearchTerm { get; set; }
        public string ResultTitle { get; set; }
        public string ReviewCount { get; set; }
        public string Rating { get; set; }
        public string ContactNumber { get; set; }
        public string Category { get; set; }
        public string Address { get; set; }
        public string StreetAddress { get; set; }
        public string City { get; set; }
        public string Zip { get; set; }
        public string Country { get; set; }
        public Dictionary<string, string> SocialMedias { get; set; }
        public string CompanyWebsite { get; set; }
        public bool Claim { get; set; }
        public string HoursInfo { get; set; }
        public Dictionary<string, string> BusinessHours { get; set; }
        public string LocatedIn { get; set; }
        public string Attributes { get; set; }
        public string MapLink { get; set; }

        public BusinessInfo(string searchTerm, string resultTitle, string reviewCount, string rating, string contactNumber,
                        string category, string address, string streetAddress, string city, string zip, string country,
                        Dictionary<string, string> socialMedias, string companyWebsite, bool claim, string hoursInfo,
                        Dictionary<string, string> businessHours, string locatedIn, string attributes, string mapLink)
        {
            SearchTerm = searchTerm;
            ResultTitle = resultTitle;
            ReviewCount = reviewCount;
            Rating = rating;
            ContactNumber = contactNumber;
            Category = category;
            Address = address;
            StreetAddress = streetAddress;
            City = city;
            Zip = zip;
            Country = country;
            SocialMedias = socialMedias ?? new Dictionary<string, string>(); // Ensure it's not null
            CompanyWebsite = companyWebsite;
            Claim = claim;
            HoursInfo = hoursInfo;
            BusinessHours = businessHours;
            LocatedIn = locatedIn;
            Attributes = attributes;
            MapLink = mapLink;
        }
    }

}
