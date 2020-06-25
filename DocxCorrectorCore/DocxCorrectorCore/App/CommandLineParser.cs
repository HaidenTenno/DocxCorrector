using System.IO;
using System.CommandLine;
using System.CommandLine.Invocation;
using DocxCorrectorCore.UserDialog;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.App
{
    public static class CommandLineParser
    {
        // Private
        /// Название параметров должно совпадать с именами при инициализации аргументов
        private static void Correct(string fileToCorrect, RulesModel rules, string paragraphsClasses, string resultPath)
        {
            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateMistakesJSON(fileToCorrect, rules, paragraphsClasses, resultPath);
        }

        private static void GoInteractive()
        {
            UserDialogCoordinator userDialogCoordinator = new UserDialogCoordinator(new QuestionsNavigationController());
            userDialogCoordinator.Start();
        }

        private static void PullProperties(string fileToAnalyse, string resultPath1, string resultPath2)
        {
            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateParagraphsPropertiesForAllTables(fileToAnalyse, resultPath1, resultPath2);
        }

        private static Command SetupCorrectCommand()
        {
            var correctCommand = new Command(name: "correct", description: "Analyze the document for formatting errors using the selected rules and class list and save the result");

            //fileToCorrect
            var fileToCorrectArg = new Argument<string>("file-to-correct");
            fileToCorrectArg.Description = "Path to the file for analysis";
            correctCommand.AddArgument(fileToCorrectArg);
            //rules
            var rulesArg = new Argument<RulesModel>("rules");
            rulesArg.Description = "Rules for verification (GOST or ITMO requirements)";
            correctCommand.AddArgument(rulesArg);
            //paragraphClasses
            var paragraphsClassesArg = new Argument<string>("paragraphs-classes");
            paragraphsClassesArg.Description = "Path to the file with paragraphs classes";
            correctCommand.AddArgument(paragraphsClassesArg);
            //resultPath
            var resultPathArg = new Argument<string>("result-path", getDefaultValue: () => Directory.GetCurrentDirectory());
            resultPathArg.Description = "File or directory path to save the result";
            correctCommand.AddArgument(resultPathArg);

            //handler
            correctCommand.Handler = CommandHandler.Create<string, RulesModel, string, string>(Correct);

            return correctCommand;
        }

        private static Command SetupPullCommand()
        {
            var pullPropertiesCommand = new Command(name: "pull", description: "Pull out the properties of document paragraphs and save them in csv");

            //fileToAnalyse
            var fileToAnalyseArg = new Argument<string>("file-to-analyse");
            fileToAnalyseArg.Description = "Path to ther file for analysis";
            pullPropertiesCommand.AddArgument(fileToAnalyseArg);
            //resultPath1
            var resultPath1Arg = new Argument<string>("result-path1", getDefaultValue: () => Directory.GetCurrentDirectory());
            resultPath1Arg.Description = "File or directory path to save the result";
            pullPropertiesCommand.AddArgument(resultPath1Arg);
            //resultPath2
            var resultPath2Arg = new Argument<string>("result-path2", getDefaultValue: () => Directory.GetCurrentDirectory());
            resultPath2Arg.Description = "File or directory path to save the result (table zero)";
            pullPropertiesCommand.AddArgument(resultPath2Arg);

            //handler
            pullPropertiesCommand.Handler = CommandHandler.Create<string, string, string>(PullProperties);

            return pullPropertiesCommand;
        }

        private static Command SetupInteractiveCommand()
        {
            var goInteractiveCommand = new Command(name: "interactive", description: "Start the program interactively");

            //handler
            goInteractiveCommand.Handler = CommandHandler.Create(GoInteractive);

            return goInteractiveCommand;
        }

        private static RootCommand SetupRootCommand()
        {
            var rootCommand = new RootCommand();

            // Correct
            var correctCommand = SetupCorrectCommand();
            rootCommand.AddCommand(correctCommand);

            // Pull paragraph properties
            var pullPropertiesCommand = SetupPullCommand();
            rootCommand.AddCommand(pullPropertiesCommand);

            // Interactive
            var goInteractiveCommand = SetupInteractiveCommand();
            rootCommand.AddCommand(goInteractiveCommand);

            return rootCommand;
        }

        // Public
        public static void Parse(string[] args)
        {
            SetupRootCommand().InvokeAsync(args).Wait();
        }
    }
}
