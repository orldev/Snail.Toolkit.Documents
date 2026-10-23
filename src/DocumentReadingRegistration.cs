using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Snail.Toolkit.Documents.Readers;

namespace Snail.Toolkit.Documents;

/// <summary>Registers document reading in a container.</summary>
public static class DocumentReadingRegistration
{
    extension(IServiceCollection services)
    {
        /// <summary>Every format reader and the <see cref="DocumentReader"/> that picks among them, as singletons.</summary>
        /// <remarks>A program that reads fewer formats registers its own readers and a <see cref="DocumentReader"/> over them instead.</remarks>
        public IServiceCollection AddDocumentReading()
        {
            foreach (var reader in FormatReaders.Every())
                services.TryAddEnumerable(ServiceDescriptor.Singleton(typeof(IFormatReader), reader));

            services.TryAddSingleton(provider => new DocumentReader(provider.GetServices<IFormatReader>(), provider.GetService<TimeProvider>()));

            return services;
        }
    }
}
