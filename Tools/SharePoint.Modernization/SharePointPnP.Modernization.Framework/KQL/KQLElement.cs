namespace SharePointPnP.Modernization.Framework.KQL
{
    /// <summary>
    /// Element in KQL query
    /// </summary>
    public class KQLElement
    {
        public string Filter { get; set; }
        public string Value { get; set; }
        public KQLFilterType Type { get; set; }
        public KQLPropertyOperator Operator { get; set; }
        public int Group { get; set; }
    }
}
