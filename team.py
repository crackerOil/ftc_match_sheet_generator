class Team:
    def __init__(self, number, name):
        self.number = number
        self.name = name
        self.matches = []
    
    def generate_matches(self, input_matches):
        self.matches = [match for match in input_matches if self in match["red_alliance"] or self in match["blue_alliance"]]