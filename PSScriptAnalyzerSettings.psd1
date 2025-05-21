@{
    # Use Severity when you want to limit the generated diagnostic records to a
    # subset of: Error, Warning and Information.
    # Uncomment the following line if you only want Errors and Warnings but
    # not Information diagnostic records.
    Severity = @('Error', 'Warning')

    # Use IncludeRules when you want to run only a subset of the default rule set.
    #IncludeRules = @('PSAvoidDefaultValueSwitchParameter',
    #                  'PSMisleadingBacktick',
    #                  'PSMissingModuleManifestField',
    #                  'PSReservedCmdletChar',
    #                  'PSReservedParams',
    #                  'PSShouldProcess',
    #                  'PSUseApprovedVerbs',
    #                  'PSUseDeclaredVarsMoreThanAssignments')

    # Use ExcludeRules when you want to run most of the default set of rules except
    # for a few rules you wish to "exclude". Note: if a rule is in both IncludeRules
    # and ExcludeRules, the rule will be excluded.
    ExcludeRules = @(
        # These rules are excluded to allow for more practical scripting patterns used in the dashcam processing scripts
        'PSUseShouldProcessForStateChangingFunctions',
        'PSAvoidUsingWriteHost',
        'PSAvoidUsingPositionalParameters'
    )

    # You can use rules from a different module or custom rules
    # by specifying the path to the module and rule name
    # CustomRulePath = 'path\to\customrule.psm1'
    # IncludeRules = @('CustomRule')
}
