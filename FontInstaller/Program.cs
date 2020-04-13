using Shell32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace FontInstaller
{
    public class Functions
    {
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: FontInstaller.exe (path to the folder of fonts to install)");
                Environment.ExitCode = 1;
                return;
            }
            var fontFolder = args[0];
            if (!Directory.Exists(fontFolder))
            {
                Console.WriteLine("Error: This directory does not exist: " + fontFolder);
                Environment.ExitCode = 2;
                return;
            }

            // Install for pre 1809 or install for 1809+
            var installVerb = GetInstallVerb();

            // Files in windows and files to be installed
            var fontFiles = GetFontFiles(Environment.ExpandEnvironmentVariables("%Windir%\\Fonts"));
            var newFontFiles = GetFontFiles(fontFolder);

            // dictionaries for current, new, new unique, and new unique installable fonts
            var fonts = GetGlyphTypefaces(fontFiles) as Dictionary<string, GlyphTypeface>;
            var newFonts = GetGlyphTypefaces(newFontFiles) as Dictionary<string, GlyphTypeface>;
            var uniqueNewFonts = GetUniques(newFonts) as Dictionary<string, GlyphTypeface>;
            var fontsToInstall = GetDifferences(fonts, uniqueNewFonts) as Dictionary<string, GlyphTypeface>;

            // Install the appropriate font files
            var totalFonts = fontsToInstall.Count();
            var counter = 1;
            foreach (var font in fontsToInstall.Keys)
            {
                Console.WriteLine("");
                Console.WriteLine("Installing file " + counter + "\\" + totalFonts + ": " + font);
                Console.WriteLine("\tFamilyName: " + String.Join(", ", fontsToInstall[font].Win32FamilyNames.Values));
                Console.WriteLine("\tFaceName[en-us]: " + fontsToInstall[font].Win32FaceNames[new System.Globalization.CultureInfo("en-us")]);
                Console.WriteLine("\tFaceNames[all]: " + String.Join(", ", fontsToInstall[font].Win32FaceNames.Values));
                InstallFont(font, installVerb);
                counter++;
            }

            Console.WriteLine("Fonts added successfully");
            Environment.ExitCode = 0;
            return;
        }

        public static string GetInstallVerb()
        {
            var installVerb = "&install";
            var os = Environment.OSVersion.Version;

            if (os.Major == 10 && os.Build >= 17763)
            {
                installVerb = "Install for &all users";
            }
            Console.WriteLine("Install verb: " + installVerb + "\n\n");
            return installVerb;
        }

        public static String[] GetFontFiles(string folder)
        {
            var ttfFiles = Directory.GetFiles(folder, "*.ttf", SearchOption.AllDirectories);
            var pfmFiles = Directory.GetFiles(folder, "*.pfm", SearchOption.AllDirectories);
            var allFiles = ttfFiles.Concat(pfmFiles).ToArray();
            Console.WriteLine("Folder: " + folder);
            Console.WriteLine("Fonts: " + allFiles.Count() + "\n\n");
            return allFiles;
        }

        public static IDictionary<string, GlyphTypeface> GetGlyphTypefaces(string[] files)
        {
            var dictionary = new Dictionary<string, GlyphTypeface>();
            foreach (var file in files)
            {
                Console.WriteLine("Added to dictionary: " + file);
                dictionary.Add(file, new GlyphTypeface(new Uri(file)));
            }
            Console.WriteLine("Total added to dictionary: " + dictionary.Count + "\n\n");
            return dictionary;
        }

        public static IDictionary<string, GlyphTypeface> GetUniques(Dictionary<string, GlyphTypeface> cullToUnique)
        {
            var uniques = new Dictionary<string, GlyphTypeface>();
            foreach (var entry in cullToUnique)
            {
                var installFont = true;
                var file = entry.Key;
                var glyph = entry.Value;
                var glyphWin32FamilyNames = glyph.Win32FamilyNames.Values.Select(s => s.ToLower()).ToArray();
                var glyphWin32FaceNames = glyph.Win32FaceNames.Values.Select(s => s.ToLower()).ToArray();
                foreach (var newFontEntry in uniques)
                {
                    var newFontFile = newFontEntry.Key;
                    var newFontGlyph = newFontEntry.Value;
                    var newFontGlyphWin32FamilyNames = newFontGlyph.Win32FamilyNames.Values.Select(s => s.ToLower()).ToArray();
                    var newFontGlyphWin32FaceNames = newFontGlyph.Win32FaceNames.Values.Select(s => s.ToLower()).ToArray();

                    // Determine if there are differences
                    var diffWin32FamilyNames = glyphWin32FamilyNames.Except(newFontGlyphWin32FamilyNames);
                    var diffWin32FaceNames = glyphWin32FaceNames.Except(newFontGlyphWin32FaceNames);
                    if (diffWin32FamilyNames.Count() == 0 && diffWin32FaceNames.Count() == 0)
                    {
                        installFont = false;
                        break;
                    }
                }
                if (installFont)
                {
                    Console.WriteLine("Unique value: " + file);
                    uniques.Add(file, glyph);
                }
            }

            // A font with a facenames array that is a superset of a duplicate font is given install priority
            // This has been imperfectly simplified to be the one with larger facenames array
            var returnDictionary = new Dictionary<string, GlyphTypeface>();
            foreach(var cursor1 in uniques)
            {
                var installFont = true;
                var file1 = cursor1.Key;
                var glyph1 = cursor1.Value;
                var glyph1Win32FamilyNames = glyph1.Win32FamilyNames.Values.Select(s => s.ToLower()).ToArray();
                var glyph1Win32FaceNames = glyph1.Win32FaceNames.Values.Select(s => s.ToLower()).ToArray();
                foreach (var cursor2 in uniques)
                {
                    var file2 = cursor2.Key;
                    var glyph2 = cursor2.Value;
                    var glyph2Win32FamilyNames = glyph2.Win32FamilyNames.Values.Select(s => s.ToLower()).ToArray();
                    var glyph2Win32FaceNames = glyph2.Win32FaceNames.Values.Select(s => s.ToLower()).ToArray();

                    var diffWin32FamilyNames = glyph1Win32FamilyNames.Except(glyph2Win32FamilyNames);
                    var diffWin32FaceNames = glyph1Win32FaceNames.Except(glyph2Win32FaceNames);
                    if (diffWin32FamilyNames.Count() == 0)
                    {
                        if (glyph1Win32FaceNames.Count() < glyph2Win32FaceNames.Count())
                        {
                            Console.WriteLine("Skipping: " + file1);
                            Console.WriteLine("\tMore FaceNames present in: " + file2);
                            installFont = false;
                            break;
                        }
                    }
                }
                if (installFont)
                {
                    returnDictionary.Add(file1, glyph1);
                }
            }

            Console.WriteLine("Total uniques: " + uniques.Count + "\n\n");
            return returnDictionary;
        }

        public static IDictionary<string, GlyphTypeface> GetDifferences(Dictionary<string, GlyphTypeface> baseFonts, Dictionary<string, GlyphTypeface> possibleFonts)
        {
            var returnDictionary = new Dictionary<string, GlyphTypeface>();
            foreach (var newFontEntry in possibleFonts)
            {
                var installFont = true;
                var newFontFile = newFontEntry.Key;
                var newFontGlyph = newFontEntry.Value;
                var newFontGlyphWin32FamilyNames = newFontGlyph.Win32FamilyNames.Values.Select(s => s.ToLower()).ToArray();
                var newFontGlyphWin32FaceNames = newFontGlyph.Win32FaceNames.Values.Select(s => s.ToLower()).ToArray();
                foreach (var entry in baseFonts)
                {
                    var file = entry.Key;
                    var glyph = entry.Value;
                    var glyphWin32FamilyNames = glyph.Win32FamilyNames.Values.Select(s => s.ToLower()).ToArray();
                    var glyphWin32FaceNames = glyph.Win32FaceNames.Values.Select(s => s.ToLower()).ToArray();

                    // Determine if there are differences
                    var diffWin32FamilyNames = glyphWin32FamilyNames.Except(newFontGlyphWin32FamilyNames);
                    var diffWin32FaceNames = glyphWin32FaceNames.Except(newFontGlyphWin32FaceNames);
                    if (diffWin32FamilyNames.Count() == 0 && diffWin32FaceNames.Count() == 0)
                    {
                        installFont = false;
                        break;
                    }
                }
                if (installFont == true)
                {
                    Console.WriteLine("Different value: " + newFontFile);
                    returnDictionary.Add(newFontFile, newFontGlyph);
                }
            }
            Console.WriteLine("Total differences: " + returnDictionary.Count + "\n\n");
            return returnDictionary;
        }

        public static string InstallFont (string file, string installVerb)
        {
            Console.WriteLine("Attempting install: " + file);
            try
            {
                var sh = (Shell32.IShellDispatch4)Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
                var folder = sh.NameSpace(Path.GetDirectoryName(file));
                var fontFile = folder.ParseName(Path.GetFileName(file));
                foreach (FolderItemVerb verb in fontFile.Verbs())
                {
                    if (verb.Name == installVerb)
                    {
                        verb.DoIt();
                        break;
                    }
                }
            }
            catch(Exception e)
            {
                return e.Message;
            }
            return "Success";
        }
    }
}
