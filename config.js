/*
 * Settings for the Dawson Wallace Subcontract Approvals add-in.
 * Fill in the three values marked SET ME once the hub site and app registration exist.
 */
window.DW_CONFIG = {
  // SET ME: your SharePoint host, e.g. "dawsonwallace.sharepoint.com"
  sharepointHost: "dawsonwallace.sharepoint.com",
  // SET ME: server-relative path of the hub site created by setup-lists.ps1
  hubSitePath: "/sites/ContractsHub",
  // SET ME: Application (client) ID of the Entra app registration "DW Subcontract Approvals"
  clientId: "00000000-0000-0000-0000-000000000000",

  graphScopes: ["Sites.ReadWrite.All", "Files.ReadWrite.All", "User.Read"],

  lists: {
    regions: "Regions",
    projects: "Projects",
    subcontracts: "Subcontracts",
    log: "Approval Log",
    requests: "Requests",
    complianceTemplate: "Compliance Template",
    complianceItems: "Compliance Items",
    settings: "App Settings"
  },

  // Word content control tags in the Appendix A templates
  tags: {
    trade: "DW_Trade",
    projectName: "DW_ProjectName",
    projectNumber: "DW_ProjectNumber",
    projectAddress: "DW_ProjectAddress",
    subcontractNumber: "DW_SubcontractNumber",
    subName: "DW_SubName",
    subAddress: "DW_SubAddress",
    signerName: "DW_SignerName",
    signerTitle: "DW_SignerTitle",
    subEmail: "DW_SubEmail",
    contractPrice: "DW_ContractPrice",
    contractPriceWords: "DW_ContractPriceWords",
    commenceDate: "DW_CommenceDate",
    substantialDate: "DW_SubstantialCompletion"
  },
  // INC checkboxes that switch on compliance items
  requirementTags: {
    bonds: "DW_Bond_INC",
    gl5m: "DW_Ins_GL5M_INC",
    gl10m: "DW_Ins_GL10M_INC",
    profLiab: "DW_Ins_ProfLiab_INC",
    wcb: "DW_WCB_INC"
  },

  // Placeholder text left in unfilled fields; treated as blank
  placeholderPatterns: [/^Enter .* Here$/i, /^\$ ?Value Here$/i, /^Written \$ Value Here$/i,
    /^Click or tap to enter a date/i, /^name@email\.com$/i, /^If none, N\/A$/i],

  // CCA 1-2021 page count (cover to signature page); Appendix A is inserted after it
  ccaPageCount: 29
};
