
=begin

Exchange 2003 ICS v1.0
LMB^Box (Thomas Montague)
Copyright (c) 2009 - 2010, LMB^Box
GNU Lesser General Public License (http://www.gnu.org/copyleft/lgpl.html)
http://labs.lmbbox.com/projects/exchange-2003-calendar-exporter/

Based off Peter Krantz's script http://www.peterkrantz.com/2006/exchange-to-ical-http/
and http://movingparts.net/2007/04/17/syncing-exchange-calendar-1-way-into-korganizerkontact/ script additions

Upgraded script to use Rexchange 0.3.4 <http://rubyforge.org/projects/rexchange/>

=end

# --- CONFIG BEGIN --- #
# Manual config section
# Change uri and options to match your settings
#icsfile = "Exchange.ics"
#uri = 'https://ExchangeServer/Exchange/Username/'
#user = 'Domain\\Username'
#password = 'Password'
#subj = "Unknown Subject"
#attendee = true
# --- CONFIG END --- #

# Defaults
icsfile = "Exchange.ics"
uri = ""
user = ""
password = ""
subj = "Unknown Subject"
attendee = true

def usage(s)
	puts("Exchange 2003 ICS v1.0")
    puts(s)
    puts("Usage: #{File.basename($0)}: [-o 'Exchange.ics'] [-r 'https://ExchangeServer/Exchange/Username/'] [-u 'Domain\\Username'] [-p 'Password'] [-s 'Unknown Subject'] [--exclude-attendee]")
    exit(2)
end

while !ARGV.empty? do
	case ARGV[0]
		when '-o':
			ARGV.shift
			icsfile = ARGV.shift
		when '-r':
			ARGV.shift
			uri = ARGV.shift
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
		when /^-/:
			usage("Unknown option: #{ARGV[0].inspect}")
		else break
	end
end

if icsfile.empty? || uri.empty? || user.empty? || password.empty? || subj.empty?
	usage("Missing Required options!")
end


# --- START --- #

puts "Processing Uri: #{uri}"

require 'rexchange'

# Transform XML-date to simple UTC format.
def xformdate(s)
	s.strftime('%Y%m%dT%H%M%S')
#	s[23..27] + s[5..6] + s[8..12] + s[14..15] + s[17..18] + "Z"
end

# We pass our uri (pointing directly to a mailbox), and options hash to RExchange::open
# to create a RExchange::Session.
RExchange::open(uri, user, password) do |mailbox|
	
	#create ics file
	calfile = File.new(icsfile, "w")
	
	calfile.puts "BEGIN:VCALENDAR"
	calfile.puts "CALSCALE:GREGORIAN"
	calfile.puts "PRODID:-//Apple Computer\, Inc//iCal 2.0//EN"
	calfile.puts "VERSION:2.0"
	
	itemcount = 0
	
	mailbox.calendar.each do |calitem|
		
		calfile.puts "BEGIN:VEVENT"
		
		# Set some properties for this event
#		calfile.puts "DTSTAMP:" + xformdate(calitem.created_at) if calitem.created_at
#		calfile.puts "DTSTART:" + xformdate(calitem.start_at) if calitem.start_at
#		calfile.puts "DTEND:" + xformdate(calitem.end_at) if calitem.end_at
#		calfile.puts "LAST-MODIFIED:" + xformdate(calitem.modified_on) if calitem.modified_on
		
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
		
#		subj = "Unknown Subject"
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
