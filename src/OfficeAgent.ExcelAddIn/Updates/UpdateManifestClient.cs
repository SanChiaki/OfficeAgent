using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateManifestClient : IUpdateManifestClient
    {
        private readonly HttpClient httpClient;

        public UpdateManifestClient(HttpClient httpClient = null)
        {
            this.httpClient = httpClient ?? new HttpClient();
        }

        public async Task<UpdateManifest> GetManifestAsync(string manifestUrl, CancellationToken cancellationToken)
        {
            if (!IsAbsoluteHttpUrl(manifestUrl))
            {
                throw new InvalidOperationException("Update manifest URL must be an absolute HTTP or HTTPS URL.");
            }

            using (var request = new HttpRequestMessage(HttpMethod.Get, manifestUrl))
            using (var response = await httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
            {
                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException($"Update manifest request failed with status code {(int)response.StatusCode} {response.StatusCode}.");
                }

                var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                UpdateManifest manifest;
                try
                {
                    manifest = JsonConvert.DeserializeObject<UpdateManifest>(body);
                }
                catch (JsonException ex)
                {
                    throw new InvalidOperationException("Update manifest JSON is invalid.", ex);
                }

                if (manifest == null || string.IsNullOrWhiteSpace(manifest.LatestVersion))
                {
                    throw new InvalidOperationException("Update manifest latest version is required.");
                }

                manifest.LatestVersion = manifest.LatestVersion.Trim();
                manifest.DownloadUrl = NormalizeHttpUrl(manifest.DownloadUrl);
                manifest.ReleaseNotesUrl = NormalizeHttpUrl(manifest.ReleaseNotesUrl);
                manifest.Title = (manifest.Title ?? string.Empty).Trim();
                manifest.Summary = (manifest.Summary ?? string.Empty).Trim();

                return manifest;
            }
        }

        private static bool IsAbsoluteHttpUrl(string value)
        {
            Uri uri;
            return Uri.TryCreate(value, UriKind.Absolute, out uri)
                && (string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));
        }

        private static string NormalizeHttpUrl(string value)
        {
            var trimmed = (value ?? string.Empty).Trim();
            return IsAbsoluteHttpUrl(trimmed) ? trimmed : string.Empty;
        }
    }
}
