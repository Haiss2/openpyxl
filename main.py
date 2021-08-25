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


wb.save("reboot_report.xlsx")
