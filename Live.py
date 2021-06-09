class Live:
    def __init__(self, influencer, date, 
    start_time, end_time, live_type, 
    weekday = None, scene = None, comment = None):
        self.influencer = influencer
        self.date = date
        self.start_time = start_time
        self.end_time = end_time
        self.live_type = live_type
        self.weekday = weekday
        self.scene = scene
        self.comment = comment
        