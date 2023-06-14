using System.ComponentModel;

namespace Capstone.Models
{
    [Serializable()]
    public class GraphicData 
    {
        [DisplayName("Hasılat")]
        public decimal Revenue { get; set; }
        [DisplayName("Satışların Maliyeti")]
        public decimal CostOfRevenue { get; set; }
        [DisplayName("Esas Faaliyetin Karı")]
        public decimal OperatingProfit { get; set; }
        [DisplayName("BRÜT KAR (ZARAR)")]
        public decimal GrossProfit { get; set; }
        [DisplayName("DÖNEM KARI (ZARARI)")]
        public decimal NetProfit { get; set; }
        public OperatingExpenses ?OperatingExpenses { get; set; }
        public Tax ?Tax { get; set; }
        [DisplayName("Amortisman ve İtfa Gideri İle İlgili Düzeltmeler")]
        public decimal Amortization { get; set; }
        public Ratios ?Ratios { get; set; }
    }
    [Serializable()]
    public class OperatingExpenses
    {
        [DisplayName("Genel Yönetim Giderleri")]
        public decimal ManagementExpenses { get; set; }
        [DisplayName("Pazarlama Giderleri")]
        public decimal MarketingExpenses { get; set; }
        [DisplayName("Araştırma ve Geliştirme Giderleri")]
        public decimal RD_Expenses { get; set; }
        [DisplayName("Esas Faaliyetlerden Diğer Gelirler")]
        public decimal OtherExpenses { get; set; }
    }
    [Serializable()]
    public class Tax
    {
        [DisplayName("Dönem Vergi (Gideri) Geliri")]
        public decimal Tax1 { get; set; }
        [DisplayName("Ertelenmiş Vergi (Gideri) Geliri")]
        public decimal Tax2 { get; set; }
    }
    [Serializable()]
    public class Ratios
    {
        public decimal GrossMargin { get; set; }
        public decimal OperatingMargin { get; set; }
        public decimal ReturnOnAssets { get; set; }
        public decimal ReturnOnEquity { get; set; }
    }
}
