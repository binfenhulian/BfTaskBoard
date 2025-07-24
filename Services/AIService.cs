using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using TaskBoard.Models;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace TaskBoard.Services
{
    public class AIService
    {
        private static readonly HttpClient httpClient = new HttpClient()
        {
            Timeout = TimeSpan.FromSeconds(30)
        };
        
        public enum AIProvider
        {
            OpenAI,
            DeepSeek,
            Claude
        }
        
        public class AIConfig
        {
            public AIProvider Provider { get; set; }
            public string ApiKey { get; set; }
        }
        
        private static string GetApiUrl(AIProvider provider)
        {
            return provider switch
            {
                AIProvider.OpenAI => "https://api.openai.com/v1/chat/completions",
                AIProvider.DeepSeek => "https://api.deepseek.com/v1/chat/completions",
                AIProvider.Claude => "https://api.anthropic.com/v1/messages",
                _ => throw new NotSupportedException($"Provider {provider} not supported")
            };
        }
        
        private static string GetModel(AIProvider provider)
        {
            return provider switch
            {
                AIProvider.OpenAI => "gpt-3.5-turbo",
                AIProvider.DeepSeek => "deepseek-chat",
                AIProvider.Claude => "claude-3-sonnet-20240229",
                _ => "gpt-3.5-turbo"
            };
        }
        
        private static void LogError(string operation, string content)
        {
            try
            {
                var logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BfTaskBoard", "ai_logs");
                Directory.CreateDirectory(logPath);
                
                var logFile = Path.Combine(logPath, $"ai_log_{DateTime.Now:yyyyMMdd}.txt");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {operation}\n{content}\n{new string('-', 80)}\n";
                
                File.AppendAllText(logFile, logEntry);
            }
            catch { }
        }

        public static async Task<TabData> GenerateTableAsync(AIConfig config, string userRequirement)
        {
            var systemPrompt = @"You are a table structure generator for BfTaskBoard. 
You MUST respond with ONLY valid JSON, no additional text or explanation.

Generate a table structure in this exact format:
{
  ""name"": ""Table Name"",
  ""columns"": [
    {
      ""id"": ""col1"",
      ""name"": ""Column Name"",
      ""type"": ""Text"",
      ""options"": []
    }
  ],
  ""sampleRow"": {
    ""col1"": ""sample value""
  }
}

Column types: Text, Single, Image, TodoList, TextArea

For Single type, include options like:
""options"": [
  {""label"": ""Option 1"", ""color"": ""#FF5722""},
  {""label"": ""Option 2"", ""color"": ""#4CAF50""}
]

CRITICAL: Return ONLY the JSON object, no markdown, no explanation, no additional text.";

            var userPrompt = $"Create a table structure for: {userRequirement}";
            
            try
            {
                LogError("REQUEST_START", $"Provider: {config.Provider}\nRequirement: {userRequirement}");
                
                string requestJson;
                StringContent content;
                
                if (config.Provider == AIProvider.Claude)
                {
                    // Claude uses different API format
                    var claudeRequest = new
                    {
                        model = GetModel(config.Provider),
                        max_tokens = 1000,
                        messages = new[]
                        {
                            new { role = "user", content = $"{systemPrompt}\n\n{userPrompt}" }
                        }
                    };
                    requestJson = JsonConvert.SerializeObject(claudeRequest);
                    content = new StringContent(requestJson, Encoding.UTF8, "application/json");
                    
                    httpClient.DefaultRequestHeaders.Clear();
                    httpClient.DefaultRequestHeaders.Add("x-api-key", config.ApiKey);
                    httpClient.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
                }
                else
                {
                    // OpenAI and DeepSeek format
                    var request = new
                    {
                        model = GetModel(config.Provider),
                        messages = new[]
                        {
                            new { role = "system", content = systemPrompt },
                            new { role = "user", content = userPrompt }
                        },
                        temperature = 0.7,
                        max_tokens = 1000
                    };
                    requestJson = JsonConvert.SerializeObject(request);
                    content = new StringContent(requestJson, Encoding.UTF8, "application/json");
                    
                    httpClient.DefaultRequestHeaders.Clear();
                    httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {config.ApiKey}");
                }
                
                LogError("REQUEST_BODY", requestJson);
                
                var response = await httpClient.PostAsync(GetApiUrl(config.Provider), content);
                var responseContent = await response.Content.ReadAsStringAsync();
                
                LogError("RESPONSE", $"Status: {response.StatusCode}\nContent: {responseContent}");
                
                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"API Error: {response.StatusCode} - {responseContent}");
                }
                
                string generatedContent;
                
                if (config.Provider == AIProvider.Claude)
                {
                    // Claude response format
                    dynamic claudeResponse;
                    try
                    {
                        claudeResponse = JsonConvert.DeserializeObject<dynamic>(responseContent);
                    }
                    catch (Exception ex)
                    {
                        LogError("PARSE_ERROR", $"Failed to parse Claude response: {ex.Message}");
                        throw new Exception($"Failed to parse Claude API response: {ex.Message}");
                    }
                    
                    if (claudeResponse == null || claudeResponse.content == null || claudeResponse.content.Count == 0)
                    {
                        throw new Exception("Invalid Claude API response format");
                    }
                    
                    generatedContent = claudeResponse.content[0].text.ToString();
                }
                else
                {
                    // OpenAI/DeepSeek response format
                    dynamic apiResponse;
                    try
                    {
                        apiResponse = JsonConvert.DeserializeObject<dynamic>(responseContent);
                    }
                    catch (Exception ex)
                    {
                        LogError("PARSE_ERROR", $"Failed to parse response: {ex.Message}");
                        throw new Exception($"Failed to parse API response: {ex.Message}");
                    }
                    
                    if (apiResponse == null || apiResponse.choices == null || apiResponse.choices.Count == 0)
                    {
                        throw new Exception("Invalid API response format");
                    }
                    
                    generatedContent = apiResponse.choices[0].message.content.ToString();
                }
                
                LogError("GENERATED_CONTENT", generatedContent);
                
                // Extract JSON from the content (in case AI added extra text)
                var jsonMatch = System.Text.RegularExpressions.Regex.Match(generatedContent, @"\{[\s\S]*\}", System.Text.RegularExpressions.RegexOptions.Multiline);
                if (!jsonMatch.Success)
                {
                    throw new Exception($"No valid JSON found in AI response. Content: {generatedContent.Substring(0, Math.Min(200, generatedContent.Length))}...");
                }
                
                // Parse the generated JSON
                dynamic tableStructure;
                try
                {
                    tableStructure = JsonConvert.DeserializeObject<dynamic>(jsonMatch.Value);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed to parse generated table structure: {ex.Message}");
                }
                
                // Convert to TabData
                var tabData = new TabData
                {
                    Id = Guid.NewGuid().ToString(),
                    Name = tableStructure.name.ToString(),
                    Columns = new List<ColumnDefinition>(),
                    Rows = new List<RowData>()
                };
                
                // Add columns
                foreach (var col in tableStructure.columns)
                {
                    var column = new ColumnDefinition
                    {
                        Id = col.id.ToString(),
                        Name = col.name.ToString(),
                        Type = Enum.Parse<ColumnType>(col.type.ToString())
                    };
                    
                    // Add options for Single type
                    if (column.Type == ColumnType.Single && col.options != null)
                    {
                        foreach (var opt in col.options)
                        {
                            column.Options.Add(new OptionItem
                            {
                                Label = opt.label.ToString(),
                                Color = opt.color.ToString()
                            });
                        }
                    }
                    
                    tabData.Columns.Add(column);
                }
                
                // Add sample row
                if (tableStructure.sampleRow != null)
                {
                    var row = new RowData
                    {
                        Id = Guid.NewGuid().ToString(),
                        Data = new Dictionary<string, object>()
                    };
                    
                    foreach (var prop in tableStructure.sampleRow)
                    {
                        row.Data[prop.Name] = prop.Value.ToString();
                    }
                    
                    tabData.Rows.Add(row);
                }
                
                return tabData;
            }
            catch (Exception ex)
            {
                LogError("ERROR", $"Exception: {ex.Message}\nStackTrace: {ex.StackTrace}");
                throw new Exception($"Failed to generate table: {ex.Message}", ex);
            }
        }
    }
}