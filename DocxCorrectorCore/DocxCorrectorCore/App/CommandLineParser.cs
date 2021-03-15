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

        private static void PullProperties(string fileToAnalyse, int initialParagraphID, string resultPath1, string resultPath2)
        {
            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateParagraphsPropertiesForAllTables(fileToAnalyse, initialParagraphID, resultPath1, resultPath2);
        }

        private static void PullWithPresets(string fileToAnalyse, string presetsFile, int initialParagraphID, string resultPath)
        {
            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateCSVWithPresetsInfo(fileToAnalyse, presetsFile, initialParagraphID, resultPath);
        }

        private static void CorrectParagraph(string fileToCorrect, RulesModel rules, int paragraphID, ParagraphClass paragraphClass, string resultPath)
        {
            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateFormattingMistakesJSON(fileToCorrect, rules, paragraphID, paragraphClass, resultPath);
        }

        private static void CreateModelFile(RulesModel rules, ParagraphClass paragraphClass, string resultPath)
        {
            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateModelJSON(rules, paragraphClass, resultPath);
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
            fileToAnalyseArg.Description = "Path to the file for analysis";
            pullPropertiesCommand.AddArgument(fileToAnalyseArg);
            //initialParagraphID
            var initialParagraphIDArg = new Argument<int>("initial-paragraph-id");
            initialParagraphIDArg.Description = "ID of the paragraph to start csv from";
            pullPropertiesCommand.AddArgument(initialParagraphIDArg);
            //resultPath1
            var resultPath1Arg = new Argument<string>("result-path1", getDefaultValue: () => Directory.GetCurrentDirectory());
            resultPath1Arg.Description = "File or directory path to save the result";
            pullPropertiesCommand.AddArgument(resultPath1Arg);
            //resultPath2
            var resultPath2Arg = new Argument<string>("result-path2", getDefaultValue: () => Directory.GetCurrentDirectory());
            resultPath2Arg.Description = "File or directory path to save the result (table zero)";
            pullPropertiesCommand.AddArgument(resultPath2Arg);

            //handler
            pullPropertiesCommand.Handler = CommandHandler.Create<string, int, string, string>(PullProperties);

            return pullPropertiesCommand;
        }

        private static Command SetupInteractiveCommand()
        {
            var goInteractiveCommand = new Command(name: "interactive", description: "Start the program interactively");

            //handler
            goInteractiveCommand.Handler = CommandHandler.Create(GoInteractive);

            return goInteractiveCommand;
        }

        private static Command SetupPullWithPresetsCommand()
        {
            var pullWithPresetsCommand = new Command(name: "pullWithPresets", description: "Pull out the properties of document paragraphs, try to classify the paragarps using presets, save result in csv");

            //fileToAnalyse
            var fileToAnalyseArg = new Argument<string>("file-to-analyse");
            fileToAnalyseArg.Description = "Path to the file for analysis";
            pullWithPresetsCommand.AddArgument(fileToAnalyseArg);
            //presetsFile
            var presetsFileArg = new Argument<string>("presets-file");
            presetsFileArg.Description = "Path to the file with presets info";
            pullWithPresetsCommand.AddArgument(presetsFileArg);
            //initialParagraphID
            var initialParagraphIDArg = new Argument<int>("initial-paragraph-id");
            initialParagraphIDArg.Description = "ID of the paragraph to start csv from";
            pullWithPresetsCommand.AddArgument(initialParagraphIDArg);
            //retultPath
            var resultPathArg = new Argument<string>("result-path", getDefaultValue: () => Directory.GetCurrentDirectory());
            resultPathArg.Description = "File or directory path to save the result";
            pullWithPresetsCommand.AddArgument(resultPathArg);

            //handler
            pullWithPresetsCommand.Handler = CommandHandler.Create<string, string, int, string>(PullWithPresets);

            return pullWithPresetsCommand;
        }

        //CorrectParagraph(string fileToCorrect, RulesModel rules, int paragraphID, ParagraphClass paragraphClass, string resultPath)
        private static Command SetupCorrectParagraphCommand()
        {
            var correctParagraphCommand = new Command(name: "correctParagraph", description: "Analyze the paragraph of the document for formatting errors using the selected rules and save the result");

            //fileToCorrect
            var fileToCorrectArg = new Argument<string>("file-to-correct");
            fileToCorrectArg.Description = "Path to the file for analysis";
            correctParagraphCommand.AddArgument(fileToCorrectArg);
            //rules
            var rulesArg = new Argument<RulesModel>("rules");
            rulesArg.Description = "Rules for verification (GOST or ITMO requirements)";
            correctParagraphCommand.AddArgument(rulesArg);
            //paragraphID
            var paragraphIDArg = new Argument<int>("paragraph-id");
            paragraphIDArg.Description = "Paragraph number in the document";
            correctParagraphCommand.AddArgument(paragraphIDArg);
            //paragraphClass
            var paragraphClassArg = new Argument<ParagraphClass>("paragraph-class");
            paragraphClassArg.Description = "Class of the selected paragraph";
            correctParagraphCommand.AddArgument(paragraphClassArg);
            //resultPath
            var resultPathArg = new Argument<string>("result-path", getDefaultValue: () => Directory.GetCurrentDirectory());
            resultPathArg.Description = "File or directory path to save the result";
            correctParagraphCommand.AddArgument(resultPathArg);

            //handler
            correctParagraphCommand.Handler = CommandHandler.Create<string, RulesModel, int, ParagraphClass, string>(CorrectParagraph);

            return correctParagraphCommand;
        }

        //CreateModelFile(RulesModel rules, ParagraphClass paragraphClass, string resultPath)
        private static Command SetupCreateModelFileCommand()
        {
            var createModelFileCommand = new Command(name: "createModelFile", description: "Get the file that contains the rules for paragraphs of a certain class for the selected requirements");

            //rules
            var rulesArg = new Argument<RulesModel>("rules");
            rulesArg.Description = "Rules for verification (GOST or ITMO requirements)";
            createModelFileCommand.AddArgument(rulesArg);
            //paragraphClass
            var paragraphClassArg = new Argument<ParagraphClass>("paragraph-class");
            paragraphClassArg.Description = "Class of the selected paragraph";
            createModelFileCommand.AddArgument(paragraphClassArg);
            //resultPath
            var resultPathArg = new Argument<string>("result-path", getDefaultValue: () => Directory.GetCurrentDirectory());
            resultPathArg.Description = "File or directory path to save the result";
            createModelFileCommand.AddArgument(resultPathArg);

            //handler
            createModelFileCommand.Handler = CommandHandler.Create<RulesModel, ParagraphClass, string>(CreateModelFile);

            return createModelFileCommand;
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

            // Pull with presets
            var pullWithPresetsCommand = SetupPullWithPresetsCommand();
            rootCommand.AddCommand(pullWithPresetsCommand);

            // Correct paragraph
            var correctParagraphCommand = SetupCorrectParagraphCommand();
            rootCommand.AddCommand(correctParagraphCommand);

            // Create model file
            var createModelFileCommand = SetupCreateModelFileCommand();
            rootCommand.AddCommand(createModelFileCommand);

            return rootCommand;
        }

        // Public
        public static void Parse(string[] args)
        {
            SetupRootCommand().InvokeAsync(args).Wait();
        }
    }
}
