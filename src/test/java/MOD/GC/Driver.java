package MOD.GC;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

public class Driver {
	
	@Test
	public static void Drive() throws InvalidFormatException, IOException
	{
		
	RES.main();
	LOB.main();
		
	}

}
