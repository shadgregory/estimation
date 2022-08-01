(ns estimation.core
  (:import [org.apache.poi.xssf.usermodel XSSFWorkbook]
           [java.io File FileOutputStream File]))

(defrecord Task [name best expected worst])

(def est-data [(Task. "Design Schema" 2 4 8)
               (Task. "Research third-pary libs" 8 10 16)
               (Task. "Web page" 16 20 32)
               (Task. "Write SQL Queries" 8 12 16)
               (Task. "REST API" 8 16 24)
               (Task. "Testing" 24 40 60)])

(defn est-struct [data]
  (loop [data data struct {}]
    (cond
      (empty? data) struct
      :else
      (recur (rest data) (assoc struct (:name (first data))
                                {:estimate (/ (+ (:best (first data))
                                                 (* (:expected (first data)) 3)
                                                 (* 2 (:worst (first data)))) 6)
                                 :stdev (/ (- (:worst (first data)) (:best (first data))) 6)
                                 :stdev2 (* 1.645 (/ (- (:worst (first data)) (:best (first data))) 6))})))))
(defn create-worksheet [the-data]
  (let  [wb (new XSSFWorkbook)
         spreadsheet (.createSheet wb)
         style (.createCellStyle wb)
         font (.createFont wb)
         header-row (.createRow spreadsheet 0)
         task-cell (.createCell header-row 0)
         est-cell (.createCell header-row 1)
         low68-cell (.createCell header-row 2)
         high68-cell (.createCell header-row 3)
         low90-cell (.createCell header-row 4)
         high90-cell (.createCell header-row 5)]
    (.setBold font true)
    (.setFont style font)
    (.setCellStyle task-cell style)
    (.setCellValue task-cell "Task")
    (.setCellStyle est-cell style)
    (.setCellValue est-cell "Estimate")
    (.setCellStyle low68-cell style)
    (.setCellValue low68-cell "Low 68")
    (.setCellStyle high68-cell style)
    (.setCellValue high68-cell "High 68")
    (.setCellStyle low90-cell style)
    (.setCellValue low90-cell "Low 90")
    (.setCellStyle high90-cell style)
    (.setCellValue high90-cell "High 90")
    (loop [data the-data cnt 1]
      (if (empty? data)
        (let [footer-row (.createRow spreadsheet cnt)
              est-total-cell (.createCell footer-row 1)]
          (.setCellValue est-total-cell (reduce
                                         #(+ %1 (double (:estimate (val %2))))
                                         0 the-data))
          wb)
        (let [row (.createRow spreadsheet cnt)]
          (.setCellValue (.createCell row 0) (key (first data)))
          (.setCellValue
           (.createCell row 1)
           (format "%.1f" (double (:estimate (val (first data))))))
          (.setCellValue (.createCell row 2) (format "%.1f" (double (- (:estimate (val (first data)))
                                                                       (:stdev (val (first data)))))))
          (.setCellValue (.createCell row 3) (format "%.1f" (double (+ (:estimate (val (first data)))
                                                                       (:stdev (val (first data)))))))
          (.setCellValue (.createCell row 4) (format "%.1f" (double (- (:estimate (val (first data)))
                                                                       (:stdev2 (val (first data)))))))
          (.setCellValue (.createCell row 5) (format "%.1f" (double (+ (:estimate (val (first data)))
                                                                       (:stdev2 (val (first data)))))))
          (recur (rest data) (inc cnt)))))))

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  (let [out (new FileOutputStream (new File "./test.xslx"))
        wb (create-worksheet (est-struct est-data))]
    (.write wb out)
    (.close out))
  (println "Hello, World!"))
