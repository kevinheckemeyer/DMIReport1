import java.util.HashMap;

public class DMIConstant {
    public enum DMIExcelCol {
        M03_Lender_Number(1),
        M04_Investor_Number(2),
        M05_Category(3),
        M07_Hi_Type(4),
        M08_Loan_Type(5),
        M10_Mortgagor_First_Name(6),
        M11_Mortgagor_Last_Name(7),
        M12_Co_Mortgagor_First_Name(8),
        M13_Co_Mortgagor_Last_Name(9),
        M14_00_Property_Address_Line_1(10),
        M14_01_Prop_Street_Dir(11),
        M15_01_Prop_Unit_No(12),
        M16_Property_City(13),
        M17_Property_State(14),
        M18_Property_Zip_Code(15),
        M19_Mortgagor_Phone_No(16),
        M20_Business_Number(17),
        M21_Employee_Code(18),
        M23_04_Mailing_Address_Line_5(19),
        M24_Mailing_Address_City(20),
        M25_Mailing_State(21),
        M26_00_Mailing_Zip_Code(22),
        M26_01_Foreign_Addr_Indic(23),
        M27_Mortgagor_SSN_TIN(24),
        M27_01_Mtg_SSN_TIN_code(25),
        M28_Co_Mortgagor_SSN_TIN(26),
        M28_01_Co_Mtg_SSN_TIN_Code(27),
        M30_01_Original_Occupancy_Status(28),
        M31_Property_Construction(29),
        M32_00_Prepayment_Penalty(30),
        M35_Original_Principal_Balance(31),
        M36_Original_Term(32),
        Closing_Settlement_Date(33),
        M40_Maturity_Date(34),
        M42_First_Payment_Date(35),
        M44_Payment_Frequency(36),
        M60_Late_Charge_Factor(37),
        M60_002_Late_Charge_Code(38),
        M60_003_Late_Charge_Min_Amt(39),
        M60_004_Late_Charge_Max_Amt(40),
        M61_Grace_Days(41),
        M62_Points_Paid_by_Borrower(42),
        M64_Investor_Loan_Number(43),
        M66_Original_Property_Value(44),
        M67_Purchase_Price(45),
        M68_Builder_Broker_Code(46),
        M69_Product_Line_Code(47),
        M72_Property_Type(48),
        M73_Loan_Purpose(49),
        Owner_Type(50),
        M75_Dist_Type(51),
        F04_Current_Principal_Balance(52),
        F05_Interest_Rate(53),
        F06_Next_Payment_Due_Date(54),
        F08_P_I_Amount_Monthly(55),
        F08_002_County_Tax(56),
        F08_003_City_Tax(57),
        F08_004_Hazard_Premium(58),
        F08_005_MIP(59),
        F08_006_Lien(60),
        F09_Escrow_Monthly_Amount(61),
        F13_Total_Payment(62),
        F15_Interest_Collected_at_Closing(63),
        F16_Escrow_Balance(64),
        F20_Accrued_Late_Charges(65),
        F21_Late_Charges_Paid_YTD(66),
        F22_Interest_Paid_YTD(67),
        F23_Principal_Paid_YTD(68),
        F24_Taxes_Paid_YTD(69),
        F29_IOE_Paid_YTD(70),
        F31_01_Effective_Yield(71),
        F32_Origination_Fees_Costs_1(72),
        F33_Origination_Fees_Costs_2(73),
        F34_Origination_Fees_Costs_3(74),
        F35_Origination_Fees_Costs_4(75),
        F36_Unamortized_Fees_Costs_1(76),
        F37_Unamortized_Fees_Costs_2(77),
        F38_Unamortized_Fees_Costs_3(78),
        F39_Unamortized_Fees_Costs_4(79),
        F42_Replacement_Reserve_Balance(80),
        Z03_Sale_Id(81),
        Z04_Man_Code(82),
        Z05_Bill_Mode(83),
        Z06_Delq_Class_Code(84),
        Z09_Billing_Cycle(85),
        Z10_Bad_Check_Stop(86),
        Z14_3_Pos_Fld_3A(87),
        Z15_Zone_Branch_Code(88),
        Z19_Cash_Batch_Code(89),
        Z20_Organization_Code(90),
        Next_tax_due_date(91),
        Z26_Branch_Code(92),
        Z30_VA_Loan_Number(93),
        Z31_RHS_Loan_number(94),
        Z32_Guaranty_Number(95),
        Z33_FHA_Number_RLT(96),
        E40_00_MORTGAGE_INSURANCE_Ins_Company(97),
        E40_01_TYPE_DISB(98),
        E40_02_SEQUENCE(99),
        E40_03_Bill_Code(100),
        E42_MORTGAGE_INSURANCE_Premium_Amt(101),
        E44_MORTGAGE_INSURANCE_Pmt_Frequency(102),
        E47_MORTGAGE_INSURANCE_Disbursement_Due_Date(103),
        E48_MORTGAGE_INSURANCE_Certificate_No(104),
        E48_02_MORTGAGE_INSURANCE_ADP_Code(105),
        X03_Mers_Min(106),
        X04_Mers_Min_Reg_Date(107),
        X05_Mers_Min_Reg_Flag(108),
        X06_Mers_Mom_Flag(109),
        A05_01_ARM_Plan_ID(110),
        A07_Lookback_Period(111),
        A08_Original_Index_Rate(112),
        A11_Original_IR_Change_Date(113),
        A12_Original_P_I_Change_Date(114),
        A13_IR_Change_Date(115),
        A14_P_I_Change_Date(116),
        A15_IR_Change_Period(117),
        A16_P_I_Change_Period(118),
        A19_Margin(119),
        A20_Single_Max_IR_increase(120),
        A22_Single_Max_IR_decrease(121),
        A24_Max_P_I_increase(122),
        A25_Max_P_I_decrease(123),
        A26_Max_Interest_Rate(124),
        A27_Min_Interest_Rate(125),
        A29_IR_Rounding_Code(126),
        A30_00_IR_Rounding_Factor(127),
        D23_Interest_Only_End_Date(128),
        Balloon_Indicator(129),
        ARM_Loan_Indicator(130),
        Higher_priced_Mortgage_Loan_Indicator(131),
        Product_Tax_Option_Code(132),
        Old_DMI_Loan_Number(133),
        MORTGAGE_INSURANCE_Upfront_Amt(134),
        Monthly_PMI_or_MI_Paid_YTD(135),
        Property_Type(136),
        Mers_ORG_ID(137),
        Original_Note_Holder_Name(138),
        Note_Funding_Date(139),
        MIN_Status_Code_Indicator(140);
        private final Integer type;
        DMIExcelCol(Integer type) {
            this.type = type;
        }
        public Integer type() {
            return type;
        }
    }
    public static Integer getLQBLoanType(String code){
        Integer returnString = null;
        HashMap<String, Integer> h = new HashMap<>();
        if (code != null) {
            h.put("Conventional",3);
            h.put("FHA",1);
            h.put("VA",2);
            h.put("USDA/Rural Housing",9);
            h.put("Other",8);

            returnString = h.get(code);
        }
        return returnString;
    }
    public static Integer getLQBOccupancyStatusType(String code){
        Integer returnString = null;
        HashMap<String, Integer> h = new HashMap<>();
        if (code != null) {
            h.put("Primary Residence",1);
            h.put("Investment",3);
            h.put("Secondary Residence",2);
            returnString = h.get(code);
        }
        return returnString;
    }
    public static String getLQBPrepaymentPenaltyType(String code){
        String returnString = "";
        HashMap<String, String> h = new HashMap<>();
        if (code != null) {
            h.put("Yes","Y");
            h.put("No","N");
            returnString = h.get(code);
        }
        return returnString;
    }
    public static Integer getLQBPropertyType(String code){
        Integer returnString = 0;
        HashMap<String, Integer> h = new HashMap<>();
        if (code != null) {
            h.put("SFR",1);
            h.put("2 Units ",4);
            h.put("3 Units ",4);
            h.put("4 Units ",4);
            h.put("Condo",3);
            h.put("Co-Op",3);
            h.put("Manufactured",5);
            h.put("Rowhouse",2);
            h.put("Modular",5);
             returnString = h.get(code);
        }
        if(returnString==null){
            returnString=0;
        }
        return returnString;
    }
    public static Integer getLQBLoanPurposeType(String code){
        Integer returnString = null;
        HashMap<String, Integer> h = new HashMap<>();
        if (code != null) {
            h.put("Purchase",1);
            h.put("Refi Rate/Term",5);
            h.put("Refinance Cash-out",6);
            h.put("Construction",4);
            h.put("Construction Perm",3);
            h.put("FHA Streamline Refi",8);
            h.put("VA IRRRL",9);
            h.put("Other",9);

            returnString = h.get(code);
        }
        return returnString;
    }
}
