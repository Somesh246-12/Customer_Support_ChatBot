import re
import random
import win32com.client as win32
import time


class chatbot:
    negative_respo = ["no","not","not a chance","not possible","never"]
    exit_commands = ["quit","pause","exit","bye","goodbye","okay","thankyou","thanks"]
    random_questions = ["With whom I'm talking?","Tell me something about yourself"]

    

    def speak(self, mssg):
        print(f"Bot: {mssg}", flush=True)
        speaker = win32.Dispatch('SAPI.SpVoice')
        speaker.Speak(mssg)
        time.sleep(1)

    def __init__(self):
        self.support_responses = {
            'ask_about_product': r'.*\bproduct\b.*',
            'technical_support': r'.*\btechnical\b.*|\bsupport\b.*',
            'about_returns': r'.*\breturn\b.*|\bpolicy\b.*',
            'general_query': r'.*\bhello\b.*|.*\bhow\b.*'
        }

    def greet(self):
        self.speak("Hello, welcome to our customer support. What's your name?")
        self.name = input("You:")

        greeting = f"Hi {self.name}.\nHow can I help you!"
        self.speak(greeting)
        will_help = input("\nYou:").lower()
        
        if will_help in self.negative_respo:
            mssg = "Alright have a great day!"
            self.speak(mssg)
            return
        self.chat() 

    def make_exit(self,reply):
        for command in self.exit_commands:
            if command in reply:
                mssg = "Thanks for reaching out, have a great day!"
                self.speak(mssg)
                return True
        return False
    
    def chat(self):
        mssg = "Please tell me your query"
        self.speak(mssg)
        reply = input("\nYou:").lower()

        while not self.make_exit(reply):
            response = self.match_reply(reply)
            self.speak(response)
            reply = input("\nYou:").lower()

    def match_reply(self,reply):
        for intent, regex_pattern in self.support_responses.items():
            found_match = re.search(regex_pattern, reply)

            if(found_match):
                if intent == 'ask_about_product':
                    return self.ask_about_product()
                
                elif intent == 'technical_support':
                    return self.technical_support()
                
                elif intent == 'about_returns':
                    return self.about_returns()
                
                elif intent == 'general_query':
                    return self.general_query()
                

        return self.no_match_intent()

    def ask_about_product(self):
        responses = ["Our product is top-notch and has excellent reviews\n","You can find out more details about the product on our website.\nWebsite link: www.wyx.in\n"]
        return random.choice(responses)
    
    def technical_support(self):
        responses = ["For technical support you can call on the number below.\nPhone: 78542xxxxx\n","You can email us at abc@gmail.com\n"]
        return random.choice(responses)
    
    def about_returns(self):
        responses = ["We have a 10-day return policy.\n","Please ensure that the product is in original condition at the time of returning.\n"]
        return random.choice(responses)
    
    def general_query(self):
        responses = ["How can I support you further?\n","Is there anything else you would like to know?\n"]
        return random.choice(responses)
    
    def no_match_intent(self):
        responses = ["I'm sorry, I didn't understand that. Can you please rephrase?\n","My apologies, can you provide ,more details?\n"]
        return random.choice(responses)
    
support_response = chatbot()
support_response.greet()