namespace SharePoint.Modernization.Scanner.Core.Telemetry
{
    /// <summary>
    /// Main telemetry events
    /// </summary>
    public enum TelemetryEvents
    {
        ScanStart,
        ScanDone,
        ScanCrash,
        GroupConnect,
        List,
        Pages,
        PublishingPortals,
        Workflows,
        InfoPath,
        Blogs,
        CustomizedForms
    }

    /// <summary>
    /// Properties that we collect for the start event
    /// </summary>
    public enum ScanStart
    {
        Mode,
        AuthenticationModel,
        SiteSelectionModel
    }

    /// <summary>
    /// Properties that we collect for the crash event
    /// </summary>
    public enum ScanCrash
    {
        ExceptionMessage,
        StackTrace
    }

    /// <summary>
    /// Properties that we collect for the done event
    /// </summary>
    public enum ScanDone
    {
        Duration
    }

    /// <summary>
    /// Measures we collect for group connect
    /// </summary>
    public enum GroupConnectResults
    {
        Sites,
        Webs,
        ReadyForGroupConnect,
        BlockingReason,
        WebTemplate,
        Warning,
        ModernUIWarning,
        PermissionWarning
    }

    /// <summary>
    /// Measures we collect for page transformation
    /// </summary>
    public enum PagesResults
    {
        Sites,
        Webs,
        Pages,
        PageType,
        PageLayout,
        IsHomePage,
        WebPartMapping,
        UnMappedWebParts
    }

    /// <summary>
    /// Measures we collect for modern list and libraries
    /// </summary>
    public enum ListResults
    {
        Sites,
        Webs,
        Lists,
        OnlyBlockedByOOB,
        RenderType,
        ListExperience,
        BaseTemplateNotWorking,
        ViewTypeNotWorking,
        MultipleWebParts,
        JSLinkWebPart,
        JSLinkField,
        XslLink,
        Xsl,
        ListCustomAction,
        PublishingField,
        SiteBlocking,
        WebBlocking,
        ListBlocking
    }

    /// <summary>
    /// Measures we collect for publishing portals
    /// </summary>
    public enum PublishingResults
    {
        Sites,
        Webs,
        Pages,
        SiteLevel,
        Complexity,
        WebTemplates,
        GlobalNavigation,
        CurrentNavigation,
        CustomSiteMasterPage,
        CustomSystemMasterPage,
        AlternateCSS,
        IncompatibleUserCustomActions,
        CustomPageLayouts,
        PageApproval,
        PageApprovalWorkflow,
        ScheduledPublishing,
        AudienceTargeting,
        Languages,
        VariationLabels,
        WebPartMapping,
        UnMappedWebParts
    }

    /// <summary>
    /// Measures collected for workflows
    /// </summary>
    public enum WorkflowResults
    {
        Workflows,
        Version,
        Scope,
        Upgradability
    }

    /// <summary>
    /// Measures collected for InfoPath
    /// </summary>
    public enum InfoPathResults
    {
        FormsFound,
        Usage
    }

    /// <summary>
    /// Measures collected for Blogs
    /// </summary>
    public enum BlogResults
    {
        Webs,
        Posts,
        Language
    }

    /// <summary>
    /// Measures collected for Customized Forms
    /// </summary>
    public enum CustomizedFormsResults
    {
        Forms,
        FormType
    }
}
