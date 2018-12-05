package com.exchange.main;

import java.util.Calendar;
import java.util.Date;

import com.exchange.utils.ExchangeUtils;

public class Main {

	public static void main(String[] args) {
		//ExchangeUtils.sendMail("mail@mail.com","Test subject", "This is a test message");
		
		Calendar c = Calendar.getInstance();
		c.set(2018, 10, 16, 10, 30, 0); //January is 0
		Date date = c.getTime();
		
		ExchangeUtils.printCalendar("maila@mail.com", date);
		
		System.out.println("Has meeting: " + ExchangeUtils.hasMeeting("mail@mail.com", date));
		
		//ExchangeUtils.createAppointment("Comité", "Se ha solicitado un comité", "Sala de consejo", new Date());
		
	}

}
