# This file contains OAPI for specific operations

def get_NumberOfStory(model):
    [StoryNumber, _, ret] = model.Story.GetNameList()
    return StoryNumber