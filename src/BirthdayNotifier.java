/**
 * Created by faster13 on 04.03.2017.
 */
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import java.io.*;

import java.net.URI;
import java.util.ArrayList;
import java.util.List;

public class BirthdayNotifier {

    // Метод загрузки настроек из файла
    private static List<String> getSettings(String pathToConfig) throws IOException {
        FileReader fr = new FileReader(pathToConfig);
        BufferedReader bf = new BufferedReader(fr);
        List<String> readLine = new ArrayList<>();
        String str;

        while ((str = bf.readLine()) != null) {
            readLine.add(str);
        }
        return readLine;
    }

    public static void main(String[] args) throws Exception {
        List<String> confSettings;

        confSettings = getSettings("src/settings.txt");

        ExchangeCredentials credentials = new WebCredentials();
        if(confSettings.size() < 2) throw new Exception();
        if(confSettings.size() == 2) credentials = new WebCredentials(confSettings.get(0), confSettings.get(1));
        if(confSettings.size() == 3) credentials = new WebCredentials(confSettings.get(0), confSettings.get(1), confSettings.get(2));
        confSettings.clear();

        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        //AutodiscoverService service = new AutodiscoverService (ExchangeVersion.Exchange2010_SP2);
        service.setCredentials(credentials);


        //service.setUrl(new URI("https://outlook.office365.com/EWS/Exchange.asmx"));
        confSettings = getSettings("src/fromto.txt");
        if(confSettings.size() != 2) throw new Exception();

        service.autodiscoverUrl(confSettings.get(0));

        EmailMessage msg= new EmailMessage(service);
        msg.setSubject("Внимание!");
        msg.setBody(MessageBody.getMessageBodyFromText("День рождения!!!"));
        msg.getToRecipients().add(confSettings.get(1));
        msg.send();

        confSettings.clear();

    }

}
