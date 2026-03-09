package io.oointerpreter.hwpcore;

import kr.dogfoot.hwp2hwpx.Hwp2Hwpx;
import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwplib.tool.textextractor.TextExtractMethod;
import kr.dogfoot.hwplib.tool.textextractor.TextExtractOption;
import kr.dogfoot.hwplib.tool.textextractor.TextExtractor;
import kr.dogfoot.hwpxlib.object.HWPXFile;
import kr.dogfoot.hwpxlib.tool.blankfilemaker.BlankFileMaker;
import kr.dogfoot.hwpxlib.writer.HWPXWriter;

import java.util.HashMap;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        try {
            if (args.length == 0) {
                throw new IllegalArgumentException("Missing command.");
            }

            String command = args[0];
            Map<String, String> options = parseOptions(args);

            switch (command) {
                case "health":
                    System.out.println("ok");
                    return;

                case "extract-hwp-text":
                    requireOption(options, "input");
                    System.out.print(extractHwpText(options.get("input")));
                    return;

                case "convert-hwp-to-hwpx":
                    requireOption(options, "input");
                    requireOption(options, "output");
                    convertHwpToHwpx(options.get("input"), options.get("output"));
                    System.out.println(options.get("output"));
                    return;

                case "create-blank-hwpx":
                    requireOption(options, "output");
                    createBlankHwpx(options.get("output"));
                    System.out.println(options.get("output"));
                    return;

                default:
                    throw new IllegalArgumentException("Unknown command: " + command);
            }
        } catch (Exception e) {
            System.err.println(e.getMessage());
            System.exit(1);
        }
    }

    private static Map<String, String> parseOptions(String[] args) {
        HashMap<String, String> options = new HashMap<String, String>();
        for (int index = 1; index < args.length; index++) {
            String arg = args[index];
            if (!arg.startsWith("--")) {
                throw new IllegalArgumentException("Unexpected argument: " + arg);
            }
            if (index + 1 >= args.length) {
                throw new IllegalArgumentException("Missing value for " + arg);
            }
            options.put(arg.substring(2), args[index + 1]);
            index += 1;
        }
        return options;
    }

    private static void requireOption(Map<String, String> options, String key) {
        if (!options.containsKey(key) || options.get(key) == null || options.get(key).trim().isEmpty()) {
            throw new IllegalArgumentException("Missing required option --" + key);
        }
    }

    private static String extractHwpText(String inputPath) throws Exception {
        HWPFile hwpFile = HWPReader.fromFile(inputPath);
        TextExtractOption option = new TextExtractOption();
        option.setMethod(TextExtractMethod.InsertControlTextBetweenParagraphText);
        option.setWithControlChar(false);
        option.setAppendEndingLF(true);
        return TextExtractor.extract(hwpFile, option);
    }

    private static void convertHwpToHwpx(String inputPath, String outputPath) throws Exception {
        HWPFile hwpFile = HWPReader.fromFile(inputPath);
        HWPXFile hwpxFile = Hwp2Hwpx.toHWPX(hwpFile);
        HWPXWriter.toFilepath(hwpxFile, outputPath);
    }

    private static void createBlankHwpx(String outputPath) throws Exception {
        HWPXWriter.toFilepath(BlankFileMaker.make(), outputPath);
    }
}
