# H/T: https://www.r-bloggers.com/using-r-to-get-data-out-of-word-docs/

library(tidyverse)
library(xml2)

get_tbls <- function(word_doc) {
  
  tmpd <- tempdir()
  tmpf <- tempfile(tmpdir=tmpd, fileext=".zip")
  
  file.copy(word_doc, tmpf)
  unzip(tmpf, exdir=sprintf("%s/docdata", tmpd))
  
  doc <- read_xml(sprintf("%s/docdata/word/document.xml", tmpd))
  
  unlink(tmpf)
  unlink(sprintf("%s/docdata", tmpd), recursive=TRUE)
  
  ns <- xml_ns(doc)
  
  tbls <- xml_find_all(doc, ".//w:tbl", ns=ns)
  
  lapply(tbls, function(tbl) {
    
    cells <- xml_find_all(tbl, "./w:tr/w:tc", ns=ns)
    rows <- xml_find_all(tbl, "./w:tr", ns=ns)
    dat <- data.frame(matrix(xml_text(cells), 
                             ncol=(length(cells)/length(rows)), 
                             byrow=TRUE), 
                      stringsAsFactors=FALSE)
    colnames(dat) <- dat[1,]
    dat <- dat[-1,]
    rownames(dat) <- NULL
    dat
    
  })
  
}

x <- get_tbls("Desktop/Opta/Opta-f24_appendices.docx")

# Table 9 is Appendix 1
x[[9]] %>%
  # Transform to tibble
  as.tibble %>% 
  # Remove the descriptions
  select(-Description) %>% 
  # The table repeats, so there are duplicate headers. filtering out rows with a
  # non-numeric `Event id` will help.
  filter(str_detect(`Event id`,"[0-9]")) %>% filter(`Event id` == 15)

# subsequent tables are Qualifiers.
x[[11]] %>% as.tibble


## install.packages("docxtractr")
library(docxtractr)
read_docx("Desktop/Opta/Opta-f24_appendices.docx") %>% docx_extract_tbl(9) %>% str
