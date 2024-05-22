package Utility;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;

public class Log {
	public static  Logger log  =  Logger.getLogger(Log.class);
	 
		public static void info(String Message)
		{
			PropertyConfigurator.configure("Log4j.properties");
			log.info(Message);
		}



}
