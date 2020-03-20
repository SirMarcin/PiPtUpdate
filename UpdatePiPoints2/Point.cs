namespace UpdatePiPoints2
{
    public class Point
    {
        public int UPW { get; set; }
        public int UPW_Kombit { get; set; }
        public string Typ_przelicznika { get; set; }
        public int WL_id { get; set; }
        public string Numer_przelicznika { get; set; }
        public string Cieplomierz_RDS_PP { get; set; }
        public string WWU1_numer { get; set; }
        public string WWU1_RDS_PP { get; set; }
        public string WWU2_numer { get; set; }
        public string WWU2_RDS_PP { get; set; }
        public string WCW_numer { get; set; }
        public string WCW_RDS_PP { get; set; }
        public string WCO_Numer { get; set; }
        public string WCO_RDS_PP { get; set; }
        public string adres_wf { get; set; }
        public string wskaznik_kontrolny { get; set; }
        public bool isChanged { get; set; }
        public bool isNew { get; set; }
    }
}
