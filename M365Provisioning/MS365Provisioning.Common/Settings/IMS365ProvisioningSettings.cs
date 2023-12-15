namespace MS365Provisioning.Common.Settings
{
    public interface IMS365ProvisioningSettings
    {
        string? GetSetting(string key);
    }
}
