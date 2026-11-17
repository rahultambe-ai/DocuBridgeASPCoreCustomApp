namespace SSL_Core_Web_App_Identify_Sample.DmsApp;

/// <summary>SharePoint or Graph answered with an error status.</summary>
public sealed class DmsUpstreamException : Exception
{
    public int StatusCode { get; }
    public string Body { get; }

    public DmsUpstreamException(int statusCode, string body)
        : base($"Upstream service answered {statusCode}.")
    {
        StatusCode = statusCode;
        Body = body;
    }
}
