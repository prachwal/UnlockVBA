using Microsoft.Extensions.CommandLineUtils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace UnlockVBA
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var app = new CommandLineApplication
                {
                    Name = "UnlockVBA"
                };
                app.HelpOption("-?|-h|--help");

                app.Command("unlock", (command) =>
                {
                    command.Description = "Decode VBA protection in Microsoft Office file";
                    command.HelpOption("-?|-h|--help");

                    var fileInOption = command.Option("-in|--FileIn", "File in path", CommandOptionType.SingleValue);
                    var fileOutOption = command.Option("-out|--FileOut", "File out path", CommandOptionType.SingleValue);

                    command.OnExecute(() =>
                    {
                        var fileIn = fileInOption.Values.FirstOrDefault();
                        var fileOut = fileOutOption.Values.FirstOrDefault();
                        var c = new MicrosoftOffice(fileIn, fileOut);
                        c.Decode();
                        return 0;
                    });
                });

                app.Command("hide", (command) =>
                {
                    command.Description = "Instruct the ninja to hide in a specific location.";
                    command.HelpOption("-?|-h|--help");

                    var locationArgument = command.Argument("[location]",
                        "Where the ninja should hide.");

                    command.OnExecute(() =>
                    {
                        var location = locationArgument.Values.Any() ?
                            String.Join(", ", locationArgument.Values.ToArray()) :
                            "under a turtle";
                        Console.WriteLine("Ninja is hidden " + location);
                        return 0;
                    });
                });

                app.Command("attack", (command) =>
                {
                    command.Description = "Instruct the ninja to go and attack!";
                    command.HelpOption("-?|-h|--help");

                    var excludeOption = command.Option("-e|--exclude <exclusions>",
                        "Things to exclude while attacking.",
                        CommandOptionType.MultipleValue);

                    var screamOption = command.Option("-s|--scream",
                        "Scream while attacking",
                        CommandOptionType.NoValue);

                    command.OnExecute(() =>
                    {
                        var exclusions = excludeOption.Values;
                        var attacking = (new List<string> {
                            "dragons",
                            "badguys",
                            "civilians",
                            "animals"
                        }).Where(x => !exclusions.Contains(x));

                        Console.Write("Ninja is attacking " + string.Join(", ", attacking));

                        if (screamOption.HasValue())
                        {
                            Console.Write(" while screaming");
                        }

                        Console.WriteLine();

                        return 0;
                    });
                });

                app.Execute(args);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}