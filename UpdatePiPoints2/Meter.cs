namespace UpdatePiPoints2
{
    public class Meter
    {
        public string GodzinyPracy { get; set; }
        public string KodBledu { get; set; }
        public string LicznikEnergii { get; set; }
        public string LicznikObjetosciPrzeplywu { get; set; }
        public string MocChwilowa { get; set; }
        public string NumerFabryczny { get; set; }
        public string PrzeplywChwilowy { get; set; }
        public string TemperaturaPowrotu { get; set; }
        public string TemperaturaZasilania { get; set; }
            
        public Meter(string Cieplomierz_RDS_PP)
        {
            GodzinyPracy = Cieplomierz_RDS_PP + ":XQ015";
            KodBledu = Cieplomierz_RDS_PP + ":XQ013";
            LicznikEnergii = Cieplomierz_RDS_PP + ":XQ001";
            LicznikObjetosciPrzeplywu = Cieplomierz_RDS_PP + ":XQ002";
            MocChwilowa = Cieplomierz_RDS_PP + ":XQ003";
            NumerFabryczny = Cieplomierz_RDS_PP + ":XQ008";
            PrzeplywChwilowy = Cieplomierz_RDS_PP + ":XQ004";
            TemperaturaPowrotu = Cieplomierz_RDS_PP + ":XQ006";
            TemperaturaZasilania = Cieplomierz_RDS_PP + ":XQ005";
        }


    }
}
