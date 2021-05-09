import nav_pages
import scrape_courses

page_1 = nav_pages.PageOne()
page_1.select_term()

page_2 = nav_pages.PageTwo()
page_2.filter_classes()