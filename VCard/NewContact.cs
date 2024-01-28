namespace ExtractAndConvert.VCard
{
    //create a new contact for each row parsed
    public class NewContact
    {
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
        public string? FormattedName { get; set; }
        public string? Organization { get; set; }
        public List<Mail>? Emails { get; set; }
        public string? Title { get; set; }
        public string? Photo { get; set; }
        public List<Phone>? Phones { get; set; }
        public List<Address>? Addresses { get; set; }
        public string? URL { get; set; }
    }

    public class Phone
    {
        public string? Number { get; set; }
    }

    public class Mail
    {
        public string? Value { get; set; }
    }

    public class Address
    {
        public string? Location { get; set; }
    }
}