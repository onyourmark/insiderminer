/* InsiderMiner, 2012. Credentials removed for publication; supply them as environment variables. */
import weka.classifiers.Evaluation;
import weka.classifiers.trees.J48;
import weka.core.Instances;
import weka.experiment.InstanceQuery;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import finance.google.*;
import java.util.ArrayList;

/*Training will be read from a table that has all the fields and U/D at the end. For first experient
	I will train on just count and U/D.
	
	Testing will be read from table tbl_keyword, but that must have the same fields (or a superset of the same fields)
	as are in the Training table.

/*

COMPILE

C:\Users\Bill\Desktop\test bed\bin>cd C:\Users\Bill\Desktop\Diamond

C:\Users\Bill\Desktop\Diamond>javac -d ./bin -cp ./lib/mail.jar src/finance/google/SendEmail
.java

C:\Users\Bill\Desktop\Diamond>
javac -d ./bin src/finance/google/PortfolioDetails.java

javac -d ./bin -cp ./lib/xmlParserAPIs-2.2.1.jar;./lib/gdata-finance-2.0.jar;./lib/gdata-core-1.0.jar;./bin src/finance/google/FinancePortfoliosClient.java


javac -d ./bin -cp ./lib/xmlParserAPIs-2.2.1.jar;./lib/gdata-finance-2.0.jar;./lib/gdata-core-1.0.jar;./lib/mail.jar;./lib/mysql-connector-java-5.1.18-bin.jar;./lib/weka.jar;./bin src/wekaJ48Coke.java


Execute:
java -cp ./lib/xmlParserAPIs-2.2.1.jar;./lib/gdata-finance-2.0.jar;./lib/gdata-core-1.0.jar;./lib/mail.jar;./lib/mysql-connector-java-5.1.18-bin.jar;./lib/weka.jar;./lib/google-collect-1.0-rc1.jar;./bin wekaJ48Coke

*/

/**
 * This example trains NaiveBayes incrementally on data obtained from the
 * ArffLoader.
 * 
 * @author FracPete (fracpete at waikato dot ac dot nz)
 */
public class wekaJ48Sony {
	
	public static String stockToGrab = "TYO:6758";
	public static String stockTrainingTable = "tbl_TrainSony";
	public static String stockTestingTable = "tbl_keywordSony";
	
	
	/**
	 * Expects an ARFF file as first argument (class attribute is assumed to be
	 * the last attribute). weather.nominalTest.arff
	 * 
	 * @param args
	 *            the commandline arguments
	 * @throws Exception
	 *             if something goes wrong
	 */
	@SuppressWarnings("unchecked")
	public static void main(String[] args) throws Exception {

		/*
		 * BufferedReader readerTrain = new BufferedReader( new
		 * FileReader("C:/Program Files/Weka-3-7/data/weather.nominal.arff"));
		 * Instances dataTrain = new Instances(readerTrain);
		 * readerTrain.close();
		 * 
		 * DataSource sourceTrain = new
		 * DataSource("C:/Program Files/Weka-3-7/data/weather.nominal.csv");
		 * Instances dataTrain = sourceTrain.getDataSet();
		 */
		 
		 // next few lines get Last Price for given stock in order to write this to the database.
		ArrayList<Object> list = new ArrayList<Object>();
		FinancePortfoliosClient cl = new FinancePortfoliosClient(
				"from-mail.com", System.getenv("PORTFOLIO_EMAIL"), System.getenv("PORTFOLIO_PASSWORD"));
		
		list = cl.getPortfolioDetails();
		if (list != null) {
			int len = list.size();
			System.out.println("LastPrice\tSymbols\tChange\tMktCap");
			for (int i = 0; i < len; i++) {
				PortfolioDetails detail = (PortfolioDetails) list.get(i);
				
				System.out.println(detail.getLastPrice() + "\t\t"
						+ detail.getSymbols() + "\t" + detail.getChange()
						+ "\t" + detail.getMktCap());
			}
		}
		// select all data from a database (I want a database that has all the
		// data fields and all the dates so I can build a model.)
		
		//this is selecting ?? all data and rebuilding the model???  Anotherwords this is the training data??
		// in that case, the database should be with all those wierd fields that Rios and I created.
		// and it will need to be updated often to rebuild the model
		InstanceQuery queryTR = new InstanceQuery();
		queryTR.setDatabaseURL(System.getenv("INSIDERMINER_DB_URL"));
		queryTR.setUsername(System.getenv("DB_READ_USER"));
		queryTR.setPassword(System.getenv("DB_READ_PASSWORD"));
		// Need a table for Training. Will call it tbl_TrainCoke
		// for initial test just use just count and U/D.
		queryTR.setQuery("select * from stockTrainingTable");
		// You can declare that your data set is sparse
		// query.setSparseData(true);
		Instances dataTrain = queryTR.retrieveInstances();
		
		//next commented out code is for reading from a local file right?

		/*
		 * BufferedReader readerTest = new BufferedReader( new
		 * FileReader("C:/Program Files/Weka-3-7/data/weather.nominalTest.arff"
		 * )); Instances dataTest = new Instances(readerTest);
		 * readerTest.close();
		 * 
		 * DataSource sourceTest = new
		 * DataSource("C:/Program Files/Weka-3-7/data/weather.nominal.csv");
		 * Instances dataTest = sourceTest.getDataSet();
		 */

		 // this next part reads from tbl_keywordCoke (stockTestingTable)
		 // I think that table needs all the same fields as training data table. I am not sure, have to check Rios video.
		 
		 // This will read from tbl_keyword but for initial test I will create a table called
		 //	tbl_TestCoke
		 // for initial test just use just count and U/D.
		 
		
		
		InstanceQuery queryTS = new InstanceQuery();
		queryTS.setDatabaseURL(System.getenv("INSIDERMINER_DB_URL"));
		queryTS.setUsername(System.getenv("DB_READ_USER"));
		queryTS.setPassword(System.getenv("DB_READ_PASSWORD"));
		
		// SET Table for Test.
		// ALERT: I don't want to read all rows so I need a WHEN = Correct Date.
		queryTS.setQuery("select * from stockTestingTable");
		
		/*
		You can declare that your data set is sparse
		query.setSparseData(true);
		*/
		// In example where I am using dataTest they are using unlabled  http://weka.wikispaces.com/Use+WEKA+in+your+Java+code
		// They use labeled but I will use dataTestLabeled
		Instances dataTest = queryTS.retrieveInstances();

		if (dataTrain.classIndex() == -1)
			dataTrain.setClassIndex(dataTrain.numAttributes() - 1);
		if (dataTest.classIndex() == -1)
			dataTest.setClassIndex(dataTest.numAttributes() - 1);
			
			
			// create copy
 Instances dataTestLabeled = new Instances(dataTest);
 
 
		// Building the model?	

		J48 cls = new J48();
		cls.buildClassifier(dataTrain);
		
		// I think that this is the right place to put this. In the example
		// they have assumed we already created the tree so here I do it after creating the tree.
		
		// Also, they have a loop here but in my case I think it will just be one record (just one day).
		
		 for (int i = 0; i < dataTest.numInstances(); i++) {
			double clsLabel = cls.classifyInstance(dataTest.instance(i));
			dataTestLabeled.instance(i).setClassValue(clsLabel);
			}
			
		// not sure if this part is right with new unlabeled data.	

		Evaluation eval = new Evaluation(dataTest);
		eval.evaluateModel(cls, dataTest);
		// starting to print to screen
		System.out.println(eval.toSummaryString("\nResults\n======\n", false));

		// just playing around with dates here?
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
		Calendar cal = Calendar.getInstance();
		// Add one day to current date time
		cal.add(Calendar.DATE, 1);
		// Add one hour to current date time
		cal.add(Calendar.HOUR, 1);

		System.out.println(dateFormat.format(cal.getTime()));
		
		// starting to write to database? To the Results table right?
		Connection con = null;
		try {
			Class.forName("com.mysql.jdbc.Driver");
			con = DriverManager.getConnection(System.getenv("INSIDERMINER_DB_URL"),
					System.getenv("DB_WRITE_USER"), System.getenv("DB_WRITE_PASSWORD"));
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
			return;
		} catch (SQLException e) {
			e.printStackTrace();
			return;
		}
//		From here should be changed because of new table.
		String sql = "INSERT INTO sampPred(var1, var2, var3, sampleText, curDate)"
				+ " values(?,?,?,?,?)";
		PreparedStatement stmt = null;
		try {
			stmt = con.prepareStatement(sql);
		} catch (SQLException e) {
			e.printStackTrace();
			if (stmt != null) {
				try {
					stmt.close();
				} catch (SQLException e1) {
					e1.printStackTrace();
					return;
				}
			}
			return;
		}

		try {
			list = cl.getPortfolioDetails();

			for (int i = 0; i < dataTest.numInstances(); i++)

			{
				double pred = cls.classifyInstance(dataTest.instance(i));
				// stmt.setString(1, dataTest.instance(i).value(0));
				stmt.setString(1, dataTest.instance(i).stringValue(0));
				
				int len = list.size();
				boolean isInserted = false;
				for (int j = 0;  j < len; j++)
				{
					PortfolioDetails detail = (PortfolioDetails) list.get(j);
					if (detail.getSymbols().equals(stockToGrab))
					{
						stmt.setString(2, detail.getLastPrice());
						isInserted = true;
					}
					
				}
				
				if (!isInserted)
					stmt.setString(2, "No exist");
				
				stmt.setString(3, dataTest.classAttribute().value((int) pred));
				stmt.setString(4, stockToGrab);
				stmt.setString(5, dateFormat.format(cal.getTime()));
				stmt.executeUpdate();
			}
		} catch (SQLException e) {
			e.printStackTrace();
			if (stmt != null) {
				try {
					stmt.close();
				} catch (SQLException e1) {
					e1.printStackTrace();
				}
			}
		}

	}
}