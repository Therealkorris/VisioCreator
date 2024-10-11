using System;
using System.Diagnostics;
using System.Text.Json.Serialization;

namespace OllamaSharp.Models
{
    public class ListModelsResponse
    {
        [JsonPropertyName("models")]
        public Model[] Models { get; set; } = null!;
    }

    [DebuggerDisplay("{Name}")]
    public class Model
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = null!;

        [JsonPropertyName("modified_at")]
        public DateTime ModifiedAt { get; set; }

        [JsonPropertyName("size")]
        public long Size { get; set; }

        [JsonPropertyName("digest")]
        public string Digest { get; set; } = null!;

        [JsonPropertyName("details")]
        public Details Details { get; set; } = null!;
    }

    public class Details
    {
        [JsonPropertyName("parent_model")]
        public string? ParentModel { get; set; }

        [JsonPropertyName("format")]
        public string Format { get; set; } = null!;

        [JsonPropertyName("family")]
        public string Family { get; set; } = null!;

        [JsonPropertyName("families")]
        public string[]? Families { get; set; }

        [JsonPropertyName("parameter_size")]
        public string ParameterSize { get; set; } = null!;

        [JsonPropertyName("quantization_level")]
        public string QuantizationLevel { get; set; } = null!;
    }
}
