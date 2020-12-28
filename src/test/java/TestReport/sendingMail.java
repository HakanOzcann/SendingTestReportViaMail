package TestReport;

import org.apache.commons.mail.DefaultAuthenticator;
import org.apache.commons.mail.EmailAttachment;
import org.apache.commons.mail.EmailException;
import org.apache.commons.mail.MultiPartEmail;
import org.testng.annotations.Test;

public class sendingMail
{

    @Test(priority = 2)
    public void sendingMail()

    {
        EmailAttachment attachment = new EmailAttachment();
        attachment.setPath("/Users/hakanozcan/Desktop/ApachePOI-Excel.xlsx");
        attachment.setDisposition(EmailAttachment.ATTACHMENT);
        attachment.setDescription("Test Excel");
        attachment.setName("Test Excel File");

        MultiPartEmail email = new MultiPartEmail();
        email.setHostName("smtp.gmail.com");
        email.setSmtpPort(465);
        email.setAuthenticator(new DefaultAuthenticator("test@gmail.com", "testpassword123"));
        email.setSSLOnConnect(true);

        try
        {
            email.setFrom("hakan.ozcan44@hotmail.com");
        }

        catch (EmailException e)
        {
            e.printStackTrace();
        }

        email.setSubject("Test Excel File");

        try
        {
            email.setMsg("This is test message.");
        }

        catch (EmailException e)
        {
            e.printStackTrace();
        }

        try
        {
            email.addTo("hakan.ozcan44@hotmail.com");
        }
        catch (EmailException e)
        {
            e.printStackTrace();
        }

        try
        {
            email.attach(attachment);
        }

        catch (EmailException e)
        {
            e.printStackTrace();
        }

        try
        {
            email.send();
        }

        catch (EmailException e)
        {
            e.printStackTrace();
        }

        System.out.println("Test file has been sent successfully.");
    }

}
