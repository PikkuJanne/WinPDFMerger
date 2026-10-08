@{
    # Safety/correctness gates for application and maintained development scripts.
    # Each category and the separate advisory pass are explained in
    # tools/test/StaticChecks.md. No severity filter or excluded/suppressed rules.
    IncludeRules = @(
        'PSAvoidAssignmentToAutomaticVariable'
        'PSAvoidDefaultValueSwitchParameter'
        'PSAvoidDefaultValueForMandatoryParameter'
        'PSAvoidUsingEmptyCatchBlock'
        'PSAvoidUsingCmdletAliases'
        'PSAvoidGlobalAliases'
        'PSAvoidGlobalFunctions'
        'PSAvoidGlobalVars'
        'PSAvoidInvokingEmptyMembers'
        'PSAvoidMultipleTypeAttributes'
        'PSAvoidNullOrEmptyHelpMessageAttribute'
        'PSAvoidOverwritingBuiltInCmdlets'
        'PSReservedCmdletChar'
        'PSReservedParams'
        'PSAvoidReservedWordsAsFunctionNames'
        'PSAvoidShouldContinueWithoutForce'
        'PSAvoidUsingUsernameAndPasswordParams'
        'PSAvoidUsingAllowUnencryptedAuthentication'
        'PSAvoidUsingBrokenHashAlgorithms'
        'PSAvoidUsingComputerNameHardcoded'
        'PSAvoidUsingConvertToSecureStringWithPlainText'
        'PSAvoidUsingDeprecatedManifestFields'
        'PSAvoidUsingInvokeExpression'
        'PSAvoidUsingPlainTextForPassword'
        'PSAvoidUsingWMICmdlet'
        'PSMisleadingBacktick'
        'PSPossibleIncorrectComparisonWithNull'
        'PSPossibleIncorrectUsageOfAssignmentOperator'
        'PSPossibleIncorrectUsageOfRedirectionOperator'
        'PSUseBOMForUnicodeEncodedFile'
        'PSUseCmdletCorrectly'
        'PSUseCompatibleSyntax'
        'PSUseConsistentParameterSetName'
        'PSUseConsistentParametersKind'
        'PSUseLiteralInitializerForHashtable'
        'PSUseProcessBlockForPipelineCommand'
        'PSUsePSCredentialType'
        'PSShouldProcess'
        'PSUseSingleValueFromPipelineParameter'
        'PSUseSupportsShouldProcess'
        'PSUseUsingScopeModifierInNewRunspaces'
    )
    Rules = @{
        PSUseCompatibleSyntax = @{
            Enable = $true
            TargetVersions = @('5.1', '7.6')
        }
    }
}
