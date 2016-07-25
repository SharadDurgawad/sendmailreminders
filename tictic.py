from twilio.rest import TwilioRestClient

accountSID = 'AC8cbd4a333c67c10d8ae9b7f4d91f0916'
authToken = '2cb23c3f21ff403b8168dde427dd540b'

twilioCli = TwilioRestClient(accountSID, authToken)

myTwilioNumber = '+19183763736'
myCellPhone = '+918056110703'

#message = twilioCli.messages.create(body="Hello Sharad", from_=myTwilioNumber, to=myCellPhone)
message = twilioCli.messages.create(to=myCellPhone, from_=myTwilioNumber, body="Hello Sharad, When are you going to Pune!")
