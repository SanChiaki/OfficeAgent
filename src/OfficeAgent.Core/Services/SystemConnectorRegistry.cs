using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public sealed class SystemConnectorRegistry : ISystemConnectorRegistry
    {
        private readonly IReadOnlyDictionary<string, ISystemConnector> connectorsBySystemKey;

        public SystemConnectorRegistry(IReadOnlyList<ISystemConnector> connectors)
        {
            if (connectors == null)
            {
                throw new ArgumentNullException(nameof(connectors));
            }

            connectorsBySystemKey = connectors
                .Where(connector => connector != null)
                .ToDictionary(
                    connector => NormalizeSystemKey(connector.SystemKey),
                    connector => connector,
                    StringComparer.OrdinalIgnoreCase);
        }

        public IReadOnlyList<ProjectOption> GetProjects()
        {
            return connectorsBySystemKey.Values
                .SelectMany(connector => connector.GetProjects() ?? Array.Empty<ProjectOption>(), NormalizeProject)
                .ToArray();
        }

        public ISystemConnector GetRequiredConnector(string systemKey)
        {
            var normalizedSystemKey = NormalizeSystemKey(systemKey);
            if (!connectorsBySystemKey.TryGetValue(normalizedSystemKey, out var connector))
            {
                throw new InvalidOperationException($"No connector is registered for systemKey '{systemKey ?? string.Empty}'.");
            }

            return connector;
        }

        private static ProjectOption NormalizeProject(ISystemConnector connector, ProjectOption project)
        {
            if (project == null)
            {
                return null;
            }

            return new ProjectOption
            {
                SystemKey = string.IsNullOrWhiteSpace(project.SystemKey) ? connector.SystemKey ?? string.Empty : project.SystemKey,
                ProjectId = project.ProjectId ?? string.Empty,
                DisplayName = project.DisplayName ?? string.Empty,
            };
        }

        private static string NormalizeSystemKey(string systemKey)
        {
            if (string.IsNullOrWhiteSpace(systemKey))
            {
                throw new InvalidOperationException("Connector systemKey is required.");
            }

            return systemKey.Trim();
        }
    }
}
