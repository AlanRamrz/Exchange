package com.exchange.utils;

import java.util.Calendar;
import java.util.Date;

import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.SendInvitationsMode;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.Mailbox;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.CalendarView;
import microsoft.exchange.webservices.data.search.FindItemsResults;

public class ExchangeUtils {

	public static void sendMail(String to, String subject, String body) {

		try {
			ExchangeService service = getService();
			EmailMessage msg = new EmailMessage(service);

			msg.setSubject(subject);
			msg.setBody(MessageBody.getMessageBodyFromText(body));
			msg.getToRecipients().add(to);
			msg.send();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static void printCalendar(String user, Date startDate) {
		Calendar c = Calendar.getInstance();
		c.setTime(startDate);
		c.add(Calendar.DAY_OF_MONTH, 1);

		Date endDate = c.getTime();
		
		try {
			ExchangeService service = getService();
			FolderId folderIdFromCalendar = new FolderId(WellKnownFolderName.Calendar,new Mailbox(user));
			CalendarFolder calendar = CalendarFolder.bind(service, folderIdFromCalendar, new PropertySet());
			CalendarView cView = new CalendarView(startDate, endDate);
			cView.setPropertySet(new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End));

			FindItemsResults<Appointment> appointments = calendar.findAppointments(cView);

			for (Appointment item : appointments) {
				System.out.println("ID: " + item.getId() + " || DATE: " + item.getStart() + " || SUBJECT: " + item.getSubject());
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	public static boolean hasMeeting(String user, Date startDate) {
		boolean res = false;

		Calendar c = Calendar.getInstance();
		c.setTime(startDate);
		c.add(Calendar.DAY_OF_MONTH, 1);

		Date endDate = c.getTime();
		
		try {
			ExchangeService service = getService();
			FolderId folderIdFromCalendar = new FolderId(WellKnownFolderName.Calendar,new Mailbox(user));
			CalendarFolder calendar = CalendarFolder.bind(service, folderIdFromCalendar, new PropertySet());
			CalendarView cView = new CalendarView(startDate, endDate);
			cView.setPropertySet(new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End));

			FindItemsResults<Appointment> appointments = calendar.findAppointments(cView);

			if(!appointments.getItems().isEmpty()) {
				res = true;
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return res;

	}
	
	public static void createAppointment(String subject, String body, String location, Date when) {
		Calendar c = Calendar.getInstance();
		c.setTime(when);
		c.add(Calendar.DAY_OF_MONTH, 1);
		
		Date startDate = c.getTime();
		
		c.add(Calendar.HOUR, 1);
		Date endDate = c.getTime();
		
		try {
			ExchangeService service = getService();
			Appointment appointment = new Appointment(service);
			
			appointment.setSubject(subject);
			appointment.setBody(new MessageBody(body));
			appointment.setLocation(location);
			appointment.setStart(startDate);
			appointment.setEnd(endDate);
			appointment.getRequiredAttendees().add("mail@mail.com");
			
			appointment.save(SendInvitationsMode.SendOnlyToAll);	
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}

	private static ExchangeService getService() throws Exception {
		ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
		ExchangeCredentials credentials = new WebCredentials("mail@mail.com", "pass");
		//ExchangeCredentials credentials = new WebCredentials("mail@mail.com", "pass");
		
		service.setCredentials(credentials);
		service.autodiscoverUrl("mail@mail.com", new RedirectionUrlCallback());
		
		return service;
	}

	static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
		public boolean autodiscoverRedirectionUrlValidationCallback(String redirectionUrl) {
			return redirectionUrl.toLowerCase().startsWith("https://");
		}
	}

}
