package Helpers;

import org.jboss.aerogear.security.otp.Totp;

public class Mf2 {

    public Mf2(){

    }
    public String getTwoFactorCode(String key){
        try{
            Totp totp = new Totp(key);
            String twoFactorCode = totp.now();
            return twoFactorCode;
        }
        catch (Exception e){
            System.out.println(String.format("Error: %s", e.getMessage()));
            return null;
        }
    }
}
