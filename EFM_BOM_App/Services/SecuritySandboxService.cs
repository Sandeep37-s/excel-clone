namespace EFM_BOM_App.Services
{
    public class SandboxExecutionResult
    {
        public bool IsSuccess { get; set; }
        public string Output { get; set; } = string.Empty;
        public string? BlockedReason { get; set; }
        public long ExecutionTimeMs { get; set; }
    }

    /// <summary>
    /// Security Governance Service.
    /// NOTE: Pattern matching, watchdog timers, and memory limits DO NOT constitute operating-system-level isolation.
    /// Because kernel-level container/AppContainer process isolation is not present in this standalone edition,
    /// dynamic execution of user-provided scripts is permanently disabled to ensure zero host vulnerability.
    /// </summary>
    public static class SecuritySandboxService
    {
        public const bool IsOsIsolationConfigured = false;
        public const bool IsDynamicScriptExecutionPermitted = false;

        public static SandboxExecutionResult ValidateAndExecute(string scriptCode, Action? scriptAction = null)
        {
            var result = new SandboxExecutionResult
            {
                IsSuccess = false,
                BlockedReason = "[SECURITY POLICY ENFORCED] Dynamic script execution is disabled. " +
                                "Operating-system-level process isolation is not implemented for standalone desktop execution. " +
                                "Arbitrary script evaluation is prohibited in this release."
            };

            LoggingService.LogWarning("Dynamic script execution attempt blocked: OS-level isolation is not configured.");
            return result;
        }
    }
}

