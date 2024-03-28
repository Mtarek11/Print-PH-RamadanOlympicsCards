using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ExcelDataReader;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Processing;
using SixLabors.Fonts;
using System.Text;
using System.Text.RegularExpressions;

public class Player
{
    public string Name { get; set; }
    public string IdNo { get; set; }
    public string Sport { get; set; }
    public string TeamName { get; set; }
    public string AgeGroup { get; set; }
}

public class Program
{
    public static async Task Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        await AddPlayerInfoToJpg();
    }

    public static async Task AddPlayerInfoToJpg()
    {
        string inputJpgPath = "Player-Card-Updated-New.jpg";
        List<Player> players = await GetAllPlayersFromExcelAsync("Padel.xlsx");

        foreach (Player player in players)
        {

            string outputJpgPath = $"Padel/{player.TeamName}_{player.Name}.jpg";

            // Load the template image
            using (var image = Image.Load(inputJpgPath))
            {
                // Create font
                var font = SystemFonts.CreateFont("Arial", 24);

                // Draw player info on the template image
                image.Mutate(ctx =>
                {
                    DrawTextIfNotNull(ctx, player.Name, font, SixLabors.ImageSharp.Color.Black, new PointF(256, 282));
                    DrawTextIfNotNull(ctx, player.IdNo, font, SixLabors.ImageSharp.Color.Black, new PointF(256, 348));
                    DrawTextIfNotNull(ctx, player.Sport, font, SixLabors.ImageSharp.Color.Black, new PointF(256, 413));
                    DrawTextIfNotNull(ctx, player.TeamName, font, SixLabors.ImageSharp.Color.Black, new PointF(256, 478));
                    DrawTextIfNotNull(ctx, player.AgeGroup, font, SixLabors.ImageSharp.Color.Black, new PointF(256, 543));
                });

                // Save the modified image as a JPG file
                image.Save(outputJpgPath, new JpegEncoder());
            }
        }
    }
    private static string SanitizeFilename(string filename)
    {
        // Remove characters that are not allowed in filenames
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            filename = filename.Replace(c, '_');
        }
        return filename;
    }

    public static async Task<List<Player>> GetAllPlayersFromExcelAsync(string filePath)
    {
        var players = new List<Player>();

        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            do
            {
                while (reader.Read())
                {
                    var playerName = GetStringOrNull(reader.GetValue(0));

                    if (!string.IsNullOrWhiteSpace(playerName))
                    {
                        //// Trim leading and trailing white spaces
                        //playerName = playerName.Trim();

                        //// Remove leading special characters like '?' and spaces
                        //playerName = playerName.TrimStart('?', ' ');

                        //// Remove extra whitespaces and non-visible characters
                        //playerName = Regex.Replace(playerName, @"\s+", " ");

                        //// Remove non-printable ASCII characters
                        //playerName = Regex.Replace(playerName, @"[^\u0020-\u007E]", "");

                        //// Ensure there are no consecutive spaces
                        //playerName = Regex.Replace(playerName, @"\s{2,}", " ");

                        //// Capitalize first letter of the name
                        //playerName = (playerName);

                        //// Create a new player instance and add it to the list
                        var player = new Player
                        {
                            Name = playerName,
                            IdNo = GetStringOrNull(reader.GetValue(1)),
                            Sport = (GetStringOrNull(reader.GetValue(2))),
                            TeamName = (GetStringOrNull(reader.GetValue(3))),
                            AgeGroup = (GetStringOrNull(reader.GetValue(4))),
                        };

                        players.Add(player);
                    }
                }
            } while (reader.NextResult());
        }

        return players;
    }
 

    private static string GetStringOrNull(object value)
    {
        return value?.ToString();
    }

    private static void DrawTextIfNotNull(IImageProcessingContext ctx, string text, Font font, Color color, PointF location)
    {
        if (!string.IsNullOrEmpty(text))
        {
            ctx.DrawText(text, font, color, location);
        }
    }
}
