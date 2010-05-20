
=begin

Exchange 2003 ICS v1.0
LMB^Box (Thomas Montague)
Copyright (c) 2009 - 2010, LMB^Box
GNU Lesser General Public License (http://www.gnu.org/copyleft/lgpl.html)
http://labs.lmbbox.com/projects/show/ex2003ics

Based off Peter Krantz's script http://www.peterkrantz.com/2006/exchange-to-ical-http/
and http://movingparts.net/2007/04/17/syncing-exchange-calendar-1-way-into-korganizerkontact/ script additions

=end

# --- CONFIG BEGIN --- #
# Manual config section
# Change URL and options to match your settings
#url = 'https://ExchangeServer/Exchange/Username/'
#user = 'Domain\\Username'
#password = 'Password'
#subj = "Unknown Subject"
#attendee = true
#icsfile = "Exchange.ics"
# --- CONFIG END --- #

# Defaults
icsfile = "Exchange.ics"
url = ""
user = ""
password = ""
subj = "Unknown Subject"
attendee = true

def usage(s)
	puts "Exchange 2003 ICS v1.0"
    puts s
    puts "Usage: #{File.basename($0)} -r 'https://ExchangeServer/Exchange/Username/' -u 'Domain\\Username' -p 'Password'"
	puts "\t\t\t[-s 'Unknown Subject'] [--exclude-attendee] [-o 'Exchange.ics']"
    exit(2)
end

while !ARGV.empty? do
	case ARGV[0]
		when '-r':
			ARGV.shift
			url = ARGV.shift
		when '-u':
			ARGV.shift
			user = ARGV.shift
		when '-p':
			ARGV.shift
			password = ARGV.shift
		when '-s':
			ARGV.shift
			subj = ARGV.shift
		when '--exclude-attendee':
			ARGV.shift
			attendee = false
		when '-o':
			ARGV.shift
			icsfile = ARGV.shift
		when /^-/:
			usage("Unknown option: #{ARGV[0].inspect}")
		else break
	end
end

if icsfile.empty? || url.empty? || user.empty? || password.empty? || subj.empty?
	usage("Missing Required options!")
end

# --- START --- #

puts "Processing URL: #{url}"

require 'rexchange'

# We pass our URL (pointing directly to a mailbox), and options hash to RExchange::open
# to create a RExchange::Session.
RExchange::open(url, user, password) do |mailbox|
	
	#create ics file
	calfile = File.new(icsfile, "w")
	
	calfile.puts "BEGIN:VCALENDAR"
	calfile.puts "CALSCALE:GREGORIAN"
	calfile.puts "PRODID:-//Apple Computer\, Inc//iCal 2.0//EN"
	calfile.puts "VERSION:2.0"
	
	itemcount = 0
	
	mailbox.calendar.each do |calitem|
	
		calfile.puts "BEGIN:VEVENT"
		
		#Set Dates with Time Zone parameter & support all_day_event
		calfile.puts "DTSTAMP;TZID=" + calitem.created_at.strftime('%Z') + ":" + calitem.created_at.strftime('%Y%m%dT%H%M%S') if calitem.created_at
		calfile.puts "LAST-MODIFIED;TZID=" + calitem.modified_on.strftime('%Z') + ":" + calitem.modified_on.strftime('%Y%m%dT%H%M%S') if calitem.modified_on
		
		if calitem.all_day_event == "1"
			calfile.puts "DTSTART:" + calitem.start_at.strftime('%Y%m%d') if calitem.start_at
			calfile.puts "DTEND:" + calitem.end_at.strftime('%Y%m%d') if calitem.end_at
		else
			calfile.puts "DTSTART;TZID=" + calitem.start_at.strftime('%Z') + ":" + calitem.start_at.strftime('%Y%m%dT%H%M%S') if calitem.start_at
			calfile.puts "DTEND;TZID=" + calitem.end_at.strftime('%Z') + ":" + calitem.end_at.strftime('%Y%m%dT%H%M%S') if calitem.end_at
		end
		
		subj = calitem.subject if calitem.subject
		calfile.puts "SUMMARY:" + subj
		
		calfile.puts "LOCATION:" + calitem.location if calitem.location
		calfile.puts "STATUS:" + calitem.meeting_status if calitem.meeting_status
		calfile.puts "DESCRIPTION:" + calitem.body.gsub("\n","\\n") if calitem.body
		calfile.puts "UID:" + calitem.uid if calitem.uid
		
		if calitem.reminder_offset
			# ouput valarm details. reminderoffset is in seconds
			ro_min = calitem.reminder_offset.to_i / 60
			calfile.puts "BEGIN:VALARM"
			calfile.puts "TRIGGER:-PT#{ro_min}M"              
			calfile.puts "DESCRIPTION:Påminnelse om aktivitet"
			calfile.puts "ACTION:DISPLAY"              
			calfile.puts "END:VALARM"
		end
		
		if attendee
			if calitem.from
				s = calitem.from
				mailto = s.gsub(/.*?<(.*?)/, '\1')
				mailto = mailto.gsub(/[<,>]/, '')
				name = s.gsub(/"(.*?), (.*?)".*/, '\2 \1')
				calfile.puts "ORGANIZER;CN=#{name}:MAILTO:#{mailto}"
			end
			
			if calitem.to
				array = calitem.to.split(">, ")
				array.each { |s| 
					mailto = s.gsub(/.*?<(.*?)/, '\1')
					mailto = mailto.gsub(/[<,>]/, '')
					name = s.gsub(/"(.*?), (.*?)".*/, '\2 \1')
					calfile.puts "ATTENDEE;CN=#{name};ROLE=REQ-PARTICIPANT:mailto:#{mailto}"
				}
			end
			
			if calitem.cc
				array = calitem.cc.split(">, ")
				array.each { |s| 
					mailto = s.gsub(/.*?<(.*?)/, '\1')
					mailto = mailto.gsub(/[<,>]/, '')
					name = s.gsub(/"(.*?), (.*?)".*/, '\2 \1')
					calfile.puts "ATTENDEE;CN=#{name};ROLE=OPT-PARTICIPANT:mailto:#{mailto}"
				}
			end
		end
		
		# add event to calendar
		calfile.puts "END:VEVENT"
		
		itemcount = itemcount + 1
		
	end
	
	# close ical file
	calfile.puts "END:VCALENDAR"
	calfile.close
	
	puts "Done! Wrote #{itemcount} appointment items"
	
end
