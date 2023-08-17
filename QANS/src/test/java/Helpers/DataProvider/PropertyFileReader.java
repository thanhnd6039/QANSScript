package Helpers.DataProvider;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Properties;

public class PropertyFileReader {
    private Properties properties;
    InputStream inputStream = null;
    public PropertyFileReader(String filePath){
        try{
            inputStream = new FileInputStream(filePath);
            properties = new Properties();
            properties.load(inputStream);
        }
        catch (Exception e){
            e.printStackTrace();
        }
        finally {
            try{
                inputStream.close();
            }
            catch (Exception e){
                e.printStackTrace();
            }
        }
    }
    public String getValueFromKey(String key){
        String value = null;
        value = properties.getProperty(key);
        return value;
    }
}
