using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Linq;
using Newtonsoft.Json;
using OfficeAgent.Infrastructure.Security;

namespace OfficeAgent.Infrastructure.Storage
{
    public sealed class FileCookieStore
    {
        private readonly string filePath;
        private readonly DpapiSecretProtector secretProtector;

        public FileCookieStore(string filePath, DpapiSecretProtector secretProtector)
        {
            this.filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            this.secretProtector = secretProtector ?? throw new ArgumentNullException(nameof(secretProtector));
        }

        public void Load(CookieContainer container)
        {
            if (container == null)
            {
                throw new ArgumentNullException(nameof(container));
            }

            if (!File.Exists(filePath))
            {
                return;
            }

            try
            {
                var encryptedContent = File.ReadAllText(filePath);
                if (string.IsNullOrWhiteSpace(encryptedContent))
                {
                    return;
                }

                string json;
                try
                {
                    json = secretProtector.Unprotect(encryptedContent);
                }
                catch (FormatException)
                {
                    return;
                }
                catch (System.Security.Cryptography.CryptographicException)
                {
                    return;
                }

                var cookies = JsonConvert.DeserializeObject<List<PersistedCookie>>(json);
                if (cookies == null)
                {
                    return;
                }

                foreach (var persisted in cookies)
                {
                    var cookie = new Cookie(
                        persisted.Name ?? string.Empty,
                        persisted.Value ?? string.Empty,
                        persisted.Path ?? "/",
                        persisted.Domain ?? string.Empty)
                    {
                        Secure = persisted.Secure,
                        HttpOnly = persisted.HttpOnly,
                    };

                    if (DateTime.TryParse(persisted.Expires, out var expires))
                    {
                        cookie.Expires = expires;
                    }

                    if (!string.IsNullOrEmpty(cookie.Name))
                    {
                        container.Add(cookie);
                    }
                }
            }
            catch (JsonException)
            {
                // Corrupt file — start fresh.
            }
            catch (IOException)
            {
                // File not accessible — start fresh.
            }
        }

        public void Save(CookieContainer container, string domain)
        {
            if (container == null)
            {
                throw new ArgumentNullException(nameof(container));
            }

            var cookies = container.GetCookies(new Uri($"https://{domain}"));
            var persisted = new List<PersistedCookie>();

            foreach (Cookie cookie in cookies)
            {
                persisted.Add(new PersistedCookie
                {
                    Name = cookie.Name,
                    Value = cookie.Value,
                    Domain = cookie.Domain,
                    Path = cookie.Path,
                    Secure = cookie.Secure,
                    HttpOnly = cookie.HttpOnly,
                    Expires = cookie.Expires.ToString("o"),
                });
            }

            var json = JsonConvert.SerializeObject(persisted, Formatting.Indented);
            var encrypted = secretProtector.Protect(json);

            var directoryPath = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            File.WriteAllText(filePath, encrypted);
        }

        public void Clear()
        {
            if (File.Exists(filePath))
            {
                try
                {
                    File.Delete(filePath);
                }
                catch (IOException)
                {
                    // Best effort — file may be locked.
                }
            }
        }

        private sealed class PersistedCookie
        {
            public string Name { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
            public string Domain { get; set; } = string.Empty;
            public string Path { get; set; } = string.Empty;
            public bool Secure { get; set; }
            public bool HttpOnly { get; set; }
            public string Expires { get; set; } = string.Empty;
        }
    }
}
