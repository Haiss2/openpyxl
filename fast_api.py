import string
import datetime as dt
from openpyxl import Workbook

from cover_page import CoverPage 
from follower_analysis import FollowerAnalysis
from m_o_m import MoM
from d_o_d import DoD
from post_list import PostList
from stories_post_list import StoriesPostList
from avg_each_type_of_post import AvgEachTypeOfPost
from top_worst_post import TopWorstPost
from top_worst_story import TopWorstStory
from ads_post_list import AdsPostList
from mention_ranking import MentionRanking
from hashtag_analysis import HashtagAnalysis
from competitor_comparision import CompetitorComparision
from competitor_top_worst  import CompetitorTopWorst
from ranked_in_hashtag import RankedInHashtag
from word_description import WordDescription
from sheet_description import SheetDescription


wb = Workbook()
std = wb.get_sheet_by_name('Sheet')
wb.remove_sheet(std)


CoverPage(wb).render()
FollowerAnalysis(wb).render()
MoM(wb).render()
DoD(wb).render()
PostList(wb).render()
AdsPostList(wb).render()
StoriesPostList(wb).render()
AvgEachTypeOfPost(wb).render()
TopWorstPost(wb).render()
TopWorstStory(wb).render()
MentionRanking(wb).render()
CompetitorComparision(wb).render()
CompetitorTopWorst(wb).render()
HashtagAnalysis(wb).render()
RankedInHashtag(wb).render()
WordDescription(wb).render()
SheetDescription(wb).render()



from fastapi import FastAPI
from fastapi.responses import FileResponse


app = FastAPI()


@app.get("/", response_class=FileResponse)
async def main():
    some_file_path = "reboot_report.xlsx"
    wb.save("reboot_report.xlsx")
    return some_file_path

