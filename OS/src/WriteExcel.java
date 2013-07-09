import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class WriteExcel {

	private static StringBuilder theFile = new StringBuilder();

	private static ArrayList<String> allTheIds = new ArrayList<String>();
	private static ArrayList<String> allTheQuestions = new ArrayList<String>();
	private static ArrayList<String> allTheAskers = new ArrayList<String>();
	private static ArrayList<Integer> medalCount = new ArrayList<Integer>();
	private static ArrayList<String> medalsFrom = new ArrayList<String>();
	private static ArrayList<String> medalsTo = new ArrayList<String>();
	private static ArrayList<String> replier = new ArrayList<String>();
	private static ArrayList<String> replyContent = new ArrayList<String>();
	private static ArrayList<Integer> numberOfReplies = new ArrayList<Integer>();

	public static void main(String[] args) throws WriteException, IOException {
		try {
			putIntoString();

			startEntry();

		} catch (IndexOutOfBoundsException e) {
			writeToSpreadSheet();
			System.out.println("done");
		}


	}

	private static void writeToSpreadSheet() throws IOException,
			RowsExceededException, WriteException {
		WritableWorkbook workbook = Workbook
				.createWorkbook(new File("lol2.xls"));
		WritableSheet sheet = workbook.createSheet("First Sheet", 0);
		Label label = new Label(0, 0, "Object Id");
		Label label2 = new Label(1, 0, "Question Content");
		Label label3 = new Label(2, 0, "User");
		Label label4 = new Label(3, 0, "# Medals");
		Label label5 = new Label(4, 0, "Medal From");
		Label label6 = new Label(5, 0, "Medal To");
		Label label7 = new Label(6, 0, "Reply From");
		Label label8 = new Label(7, 0, "Reply Content");
		Label label9 = new Label(8, 0, "# Replies");

		sheet.addCell(label);
		sheet.addCell(label2);
		sheet.addCell(label3);
		sheet.addCell(label4);
		sheet.addCell(label5);
		sheet.addCell(label6);
		sheet.addCell(label7);
		sheet.addCell(label8);
		sheet.addCell(label9);

		int medalRowCounter = 1;
		int medalToRowCounter = 1;
		int idCounter = 0;
		int rowCounter = 1;
		int medalFromIndex = 0;
		int medalToIndex = 0;
		int replyFromIndex = 0;
		int replyContentIndex = 0;
		int replyFromRowCounter = 1;
		int replyContentRowCounter = 1;
		int largestCellLengthSoFar = 0;

		for (int i = 0; i < allTheIds.size(); i++) {

			Label id = new Label(0, rowCounter, allTheIds.get(idCounter));
			sheet.addCell(id);
			Label question = new Label(1, rowCounter,
					allTheQuestions.get(idCounter));
			sheet.addCell(question);

			Label user = new Label(2, rowCounter, allTheAskers.get(idCounter));
			sheet.addCell(user);


			Label numberOfMedals = new Label(3, rowCounter, medalCount.get(
					idCounter).toString());
			sheet.addCell(numberOfMedals);
			Label allTheNumbersOfReplies = new Label(8, rowCounter,
					numberOfReplies.get(idCounter).toString());
			sheet.addCell(allTheNumbersOfReplies);

			int medalLength = medalCount.get(idCounter);

			if (medalLength == 0) {
				medalLength = 1;
			}

			try {
				for (int k = 0; k < medalLength; k++) {

					Label medalFrom = new Label(4, medalRowCounter,
							medalsFrom.get(medalFromIndex));
					sheet.addCell(medalFrom);
					medalRowCounter++;
					medalFromIndex++;

				}


				for (int l = 0; l < medalLength; l++) {
					Label medalTo = new Label(5, medalToRowCounter,
							medalsTo.get(medalToIndex));
					sheet.addCell(medalTo);
					medalToRowCounter++;
					medalToIndex++;

				}

				int replyLength = numberOfReplies.get(idCounter);
				int copyOfReplyFromIndex = replyFromIndex;

				for (int m = 0; m <= replyLength; m++) {

					Label replyFrom = new Label(6, replyFromRowCounter,
							replier.get(replyFromIndex));
					sheet.addCell(replyFrom);
					replyFromRowCounter++;
					replyFromIndex++;

				}
				//dsa
				if (copyOfReplyFromIndex == replyFromIndex -1) {
					;
				} else {

					replyFromIndex = replyFromIndex - 1;
				}
				int copyOfContentIndex = replyContentIndex;

				for (int n = 0; n <= replyLength; n++) {

					Label theReply = new Label(7, replyContentRowCounter,
							replyContent.get(replyContentIndex));
					sheet.addCell(theReply);
					replyContentRowCounter++;
					replyContentIndex++;

				}
				if (copyOfContentIndex == replyContentIndex - 1) {
					;
				} else {

					replyContentIndex = replyContentIndex - 1;
				}
			} catch (IndexOutOfBoundsException e) {
				System.out.println("k");
			}

			largestCellLengthSoFar = medalCount.get(idCounter);

			if (largestCellLengthSoFar >= numberOfReplies.get(idCounter)) {

				if (largestCellLengthSoFar == 0) {
					rowCounter += 1;
				} else {
					rowCounter += largestCellLengthSoFar;
					medalRowCounter = rowCounter;
					medalToRowCounter = rowCounter;
					replyFromRowCounter = rowCounter;
					replyContentRowCounter = rowCounter;
				}

			}

			else {


				rowCounter += numberOfReplies.get(idCounter);
				medalRowCounter = rowCounter;
				medalToRowCounter = rowCounter;
				replyFromRowCounter = rowCounter;
				replyContentRowCounter = rowCounter;

			}

			idCounter++;

		}

		workbook.write();
		workbook.close();

	}

	private static void startEntry() {
		String idTemp = "";

		int idLocation;

		while (theFile.indexOf("\"_id\"") > 0) {
		

			idLocation = theFile.indexOf("\"_id\"") + 18;

			idTemp = theFile.substring(idLocation, idLocation + 24);

			System.out.println(idTemp.toString());
			allTheIds.add(idTemp);
			theFile = theFile.delete(0, idLocation + 24);

			getQuestion();

		}

	}

	private static void getQuestion() {

		String idTemp = "";

		int questionLocation;
		int locationOfQuestionEnd;
		locationOfQuestionEnd = theFile.indexOf("\",");
		questionLocation = theFile.indexOf("\"body\"") + 10;

		idTemp = theFile.substring(questionLocation, locationOfQuestionEnd);
		System.out.println(idTemp.toString());
		allTheQuestions.add(idTemp);
		theFile = theFile.delete(0, locationOfQuestionEnd + 2);

		getAsker();

	}

	private static void getAsker() {
		String idTemp = "";

		int askerLocation;
		int locationOfAskerEnd;

		askerLocation = theFile.indexOf("\"from\"") + 10;
		locationOfAskerEnd = theFile.indexOf("\",");

		idTemp = theFile.substring(askerLocation, locationOfAskerEnd);

		System.out.println(idTemp.toString());
		allTheAskers.add(idTemp);
		theFile = theFile.delete(0, locationOfAskerEnd + 2);

		getMedals();
	}

	private static void getMedals() {

		String idTemp = "";

		int medalFromLocation;
		int locationOfMedalFromEnd;
		int locationOfMedalToEnd;
		medalFromLocation = theFile.indexOf("\"medals\"") + 11;
		String checkMedal = theFile.substring(medalFromLocation + 2,
				medalFromLocation + 4);
		if (checkMedal.equals("],")) {
			medalCount.add(0);
			medalsFrom.add("N/A");
			medalsTo.add("N/A");

		}

		else {
			String doneWithMedals = "},";
			int medalsCount = 0;

			while (doneWithMedals.equals("},")) {

				int medalFrom;
				medalFrom = theFile.indexOf("\"from\"") + 10;
				locationOfMedalFromEnd = theFile.indexOf("\",");
				idTemp = theFile.substring(medalFrom, locationOfMedalFromEnd);
				System.out.println(idTemp.toString());
				medalsFrom.add(idTemp);
				theFile = theFile.delete(0, locationOfMedalFromEnd + 2);

				int medalTo;
				medalTo = theFile.indexOf("\"to\"") + 8;
				locationOfMedalToEnd = theFile.indexOf("\",");
				idTemp = theFile.substring(medalTo, locationOfMedalToEnd);
				System.out.println(idTemp.toString());
				medalsTo.add(idTemp);
				theFile = theFile.delete(0, locationOfMedalToEnd + 2);
				doneWithMedals = theFile.substring(locationOfMedalToEnd + 34,
						locationOfMedalToEnd + 36);
				medalsCount++;

			}
			medalCount.add(medalsCount);
			theFile = theFile.delete(0, theFile.indexOf("\"replies\"") - 1);

		}

		getReplies();

	}

	private static void getReplies() {

		String idTemp = "";

		int fromLocation;
		int locationOfFromEnd;
		int locationOfReplyEnd;
		fromLocation = theFile.indexOf("\"replies\"") + 12;
		String checkReply = theFile.substring(fromLocation + 2,
				fromLocation + 4);
		if (checkReply.equals("]\n")) {
			numberOfReplies.add(0);
			replier.add("N/A");
			replyContent.add("N/A");

		}

		else {
			String doneWithReplies = "{\n";
			int replyCount = 0;

			while (doneWithReplies.contains("{\n")
					&& !(doneWithReplies.contains("]\n"))) {

				int replyFrom;
				replyFrom = theFile.indexOf("\"from\"") + 10;
				locationOfFromEnd = theFile.indexOf("\",");
				idTemp = theFile.substring(replyFrom, locationOfFromEnd);
				System.out.println(idTemp.toString());
				replier.add(idTemp);
				theFile = theFile.delete(0, locationOfFromEnd + 2);

				int reply;
				reply = theFile.indexOf("\"body\"") + 10;
				locationOfReplyEnd = theFile.indexOf("\",\n");
				idTemp = theFile.substring(reply, locationOfReplyEnd);
				System.out.println(idTemp.toString());
				replyContent.add(idTemp);
				theFile = theFile.delete(0, locationOfReplyEnd + 2);
		
				try {
					doneWithReplies = theFile.substring(156, 325);
				} catch (IndexOutOfBoundsException e) {

					numberOfReplies.add(replyCount + 1);
				}
	
				theFile = theFile.delete(0, theFile.indexOf("{\n"));
				replyCount++;

			}

			numberOfReplies.add(replyCount);



		}
	}

	private static void putIntoString() throws FileNotFoundException {
		File input = new File("openstudy.txt");
		Scanner scan = new Scanner(input);

		while (scan.hasNext()) {
			theFile.append(scan.nextLine());
			theFile.append("\n");

		}

	}

}
