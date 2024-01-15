namespace EnvioInforme
{
    public class OperationRequest
    {
        public string? UserName { get; set; }
        public string? ReportName { get; set; }
        public string? ReportGuid { get; set; }
        public string? ReportType { get; set; }
        public string? ReportTittle { get; set; }
        public int TotalImage { get; set; }

        #region Report RFDC
        public string? RfdcOriginRefrigerator { get; set; }
        public string? RfdcDispatchType { get; set; }
        public string? RfdcCustomer { get; set; }
        public string? RfdcDestination { get; set; }
        public string? RfdcSpecie { get; set; }
        public string? RfdcStatus { get; set; }
        public string? RfdcObservation { get; set; }
        #endregion

        #region Report RFC
        public string? RfcDescription { get; set; }
        public string? RfcObservation { get; set; }
        #endregion

        #region Report RFF
        public string? RffType { get; set; }
        public string? RffDispatchDate { get; set; }
        public string? RffDestination { get; set; }
        public string? RffSpecie { get; set; }
        public string? RffStatus { get; set; }
        public string? RffPortDeparture { get; set; }
        public string? RffPortDestination { get; set; }
        public string? RffObservation { get; set; }
        #endregion

    }

    public class OperationResponse
    {
        public bool OperationResult { get; set; }
    }

    public class ImageRequest
    {
        public string? ImageName { get; set; }
        public string? ImageB64 { get; set; }
    }
}