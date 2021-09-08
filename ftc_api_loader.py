import requests
import base64

from team import Team

class FTCApiLoader:
    def __init__(self, event_url, api_token):
        self.event_url = event_url
        self.api_token = api_token

    def load_teams(self):
        response = requests.get(
            "http://ftc-api.firstinspires.org/v2.0/2020/teams?eventCode={}".format(self.event_url.split("/")[-1]),
            headers={"Authorization": "Basic {}".format(base64.b64encode(self.api_token.encode("ascii")).decode("ascii"))}
        )

        if self.check_response_status(response):
            teams = []
            for team in response.json()["teams"]:
                teams.append(Team(team["teamNumber"], team["nameShort"]))

            return teams

        return []

    def load_matches(self, teams):
        response = requests.get(
            "http://ftc-api.firstinspires.org/v2.0/2020/schedule/{}?tournamentLevel=qual".format(self.event_url.split("/")[-1]),
            headers={"Authorization": "Basic {}".format(base64.b64encode(self.api_token.encode("ascii")).decode("ascii"))}
        )

        if self.check_response_status(response):
            matches = []
            for match in response.json()["schedule"]:
                matches.append({
                    "name": match["description"], 
                    "red_alliance": [next(team for team in teams if team.number == match["teams"][0]["teamNumber"]), next(team for team in teams if team.number == match["teams"][1]["teamNumber"])],
                    "blue_alliance": [next(team for team in teams if team.number == match["teams"][2]["teamNumber"]), next(team for team in teams if team.number == match["teams"][3]["teamNumber"])]
                })
            
            return matches

        return []

    def check_response_status(self, response):
        if response.status_code == 400:
            print("HTTP 400 - Invalid Season Requested/Malformed Parameter Format In Request/Missing Parameter In Request/Invalid API Version Requested")
        elif response.status_code == 401:
            print("HTTP 200 - Unauthorized")
        elif response.status_code == 404:
            print("HTTP 404 - Invalid Event Requested")
        elif response.status_code == 500:
            print("HTTP 500 - Internal Server Error")
        elif response.status_code == 501:
            print("HTTP 501 - Request Did Not Match Any Current API Pattern")
        elif response.status_code == 503:
            print("HTTP 503 - Service Unavailable")
        else:
            return True

        return False
