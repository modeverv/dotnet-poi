using System;
Console.WriteLine($"AppContext.BaseDirectory = {AppContext.BaseDirectory}");
Console.WriteLine($"Directory.GetCurrentDirectory() = {Directory.GetCurrentDirectory()}");
Console.WriteLine($"Environment.CurrentDirectory = {Environment.CurrentDirectory}");
Console.WriteLine($"poi/test-data exists: {Directory.Exists(Path.GetFullPath("poi/test-data"))}");
Console.WriteLine($"poi/test-data from cwd: {Directory.Exists(Path.Combine(Directory.GetCurrentDirectory(), "poi/test-data"))}");
