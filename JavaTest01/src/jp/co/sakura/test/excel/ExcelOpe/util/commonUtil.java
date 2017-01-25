package jp.co.sakura.test.excel.ExcelOpe.util;

import java.util.Hashtable;

import javax.naming.AuthenticationException;
import javax.naming.Context;
import javax.naming.NamingException;
import javax.naming.directory.DirContext;
import javax.naming.directory.InitialDirContext;

import org.apache.log4j.Logger;

import jp.co.sakura.test.excel.ExcelOpe.dto.ExcelProp;

public class commonUtil {

	private static Logger log = Logger.getLogger(commonUtil.class);
	private static DirContext ctx = null;

	// 全角数字を半角数字に変換する
	private static final String[] numeric = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };

	 private static  String numTohan(char ch){
	    	if ( ch >= '０' && ch <= '９' ) {
	    		return numeric[ ch - '０' ];
	    	} else {
	    		return String.valueOf( ch );
	    	}
	    }

    public static String exchangeSmall(String s) {

    	StringBuffer buffer = new StringBuffer();
    	for (int i = 0; i < s.length(); i++ ) {
    		char ch = s.charAt(i);
    		buffer.append(numTohan(ch));
    	}
    	return zeroPudding(buffer.toString());
    }

    public static String zeroPudding(String s) {

    	if(s.equals("1")) {
    		s = "01";
    	} else if(s.equals("2")) {
    		s = "02";
       	} else if(s.equals("3")) {
       		s = "03";
       	} else if(s.equals("4")) {
       		s = "04";
       	} else if(s.equals("5")) {
       		s = "05";
       	} else if(s.equals("6")) {
       		s = "06";
       	} else if(s.equals("7")) {
       		s = "07";
       	} else if(s.equals("8")) {
       		s = "08";
       	} else if(s.equals("9")) {
       		s = "09";
       	}
    	return s;
    }

    public static int ActiveDirectoryAuth() {

    	int iRtn = 0;

		Hashtable<String,String> env = new Hashtable<String,String>();
		env.put(Context.INITIAL_CONTEXT_FACTORY,"com.sun.jndi.ldap.LdapCtxFactory");
		env.put(Context.PROVIDER_URL, ExcelProp.getLdapPath());
		env.put(Context.SECURITY_AUTHENTICATION, "simple");
		env.put(Context.SECURITY_PRINCIPAL, ExcelProp.getAdUser() + "@" + ExcelProp.getDomainName());
		env.put(Context.SECURITY_CREDENTIALS, ExcelProp.getAdPass());
		try {
			// bind認証する
			ctx = new InitialDirContext(env);

			log.info("認証OK");
		} catch (AuthenticationException ae) {
			log.info("認証NG");
			iRtn = 1;
		} catch (Exception e) {
			log.info("エラー");
			iRtn = 1;
		}
		return iRtn;
    }

    public static void ActiveDirectoryAuthClose() {
    	try {
			ctx.close();
		} catch (NamingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }

}
