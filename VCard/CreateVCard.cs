using System.Text;

namespace ExtractAndConvert.VCard
{
    public static class CreateVCard
    {
        public static string VCard(NewContact contact)
        {
            StringBuilder vcard = new StringBuilder();
            vcard.Append("BEGIN:VCARD");
            vcard.Append("\nVERSION:2.1");

            //Name
            vcard.Append($"\nN:{contact.LastName};{contact.FirstName};");

            //Full Name
            vcard.Append($"\nFN:{contact.FirstName} {contact.LastName}");

            //Organization name
            vcard.Append($"\nORG:{contact.Organization}");

            //Title
            vcard.Append($"\nTITLE:{contact.Title}");

            //Numbers
            if (contact.Phones != null && contact.Phones.Count > 0)
            {
                foreach (var phoneNumber in contact.Phones)
                {
                    vcard.Append($"\nTEL;WORK;VOICE:{phoneNumber.Number}");
                }
            }

            //Addresses
            if (contact.Addresses != null && contact.Addresses.Count > 0)
            {
                foreach (var address in contact.Addresses)
                {
                    vcard.Append($"\nADR;WORK,PREF:;;{address.Location}");
                }
            }

            //Email
            if (contact.Emails != null && contact.Emails.Count > 0)
            {
                foreach (var email in contact.Emails)
                {
                    vcard.Append($"\nEMAIL:{email.Value}");
                }
            }

            //Links
            vcard.Append($"\nURL:{contact.URL}");

            //end of vcard
            vcard.Append("\nEND:VCARD");

            return vcard.ToString();
        }
    }
}