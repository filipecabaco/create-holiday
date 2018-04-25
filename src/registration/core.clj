(ns registration.core
  (:gen-class)
  (:require [dk.ative.docjure.spreadsheet :as ss]
            [clojure.java.io :as io]
            [clojure.string :as string])
  (:import [org.apache.commons.io IOUtils]
           [java.time LocalDate]
           [org.apache.poi.ss.usermodel WorkbookFactory]
           [org.apache.poi.ss.usermodel CellType]))

(def months '("Jan" "Feb" "Mar" "Apr" "May" "Jun" "Jul" "Aug" "Sep" "Oct" "Nov" "Dec"))

;; BEGIN READ VACATION SHEET
(defn- get-month-number [month]
  (->> months
       (filter #(= month %1))
       first
       (.indexOf months)
       inc))

(defn- is-number [cell]
  (try
    (not (zero? (.getNumericCellValue cell)))
    (catch Exception e false)))

(defn- get-days [row]
  (let [all                     (reverse (iterator-seq (.cellIterator row)))
        remove-filler           (drop-while (complement is-number) all)
        {days true empty false} (group-by is-number remove-filler)
        empty-count             (count empty)
        days-count              (count days)]
    {:empty empty-count :days days-count}))

(defn- filter-by-name [name row]
  (let [cell     (.getCell row 0)
        cellName (.getStringCellValue cell)]
    (= cellName name)))

(defn- process-cell [cell]
  (let [value (.getStringCellValue cell)]
    (cond
      (re-matches #"(?i).*AL.*" value) :vacation
      :else                            :full)))

(defn- process-row [skip month-days month row]
  (let [all          (iterator-seq (.cellIterator row))
        [name-cell]  all
        name         (.getStringCellValue name-cell)
        skipped      (drop skip all)
        cells        (take month-days skipped)
        days         (if-not (empty? name) (seq (map process-cell cells)) nil)
        month_number (get-month-number month)]
    (if-not (nil? days)
      {:name name :days days :month month_number})))

(defn- get-rows-from-sheet
  [sheet]
  (let [last-row (.getLastRowNum sheet)
        rows     (map #(.getRow sheet %1) (range last-row))]
    rows))

(defn read-sheet [name workbook month]
  (let [sheet                     (.getSheet workbook month)
        rows                      (get-rows-from-sheet sheet)
        [month-row & rows]        (drop 4 rows)
        {empty :empty days :days} (get-days month-row)
        selected                  (filter (partial filter-by-name name) rows)]
    (map (partial process-row empty days month) selected)))

(defn read-vacation [name file]
  (with-open [workbook (WorkbookFactory/create (io/file file))]
    (map (partial read-sheet name workbook) months)))

(defn read-vacations [name file]
  (read-vacation name file))

;; END READ VACATION SHEET

;; START CREATE EVENT
(defn is-full [all]
  (let [{status :status} all]
    (= status :full)))

(defn create-day [year month [day status]]
  {:year   year
   month   :month
   :day    (inc day)
   :status status})

(defn create-event [year schedule]
  (let [{days :days month :month} schedule
        days-with-index           (map-indexed vector days)
        status-map                (map (partial create-day year month) days-with-index)
        filtered-status-map       (filter (complement is-full) status-map)]
    (println filtered-status-map)))

(defn create-events [year schedule]
  (doall (map (partial create-event year) schedule)))

;; END CREATE EVENT

;; START PROCESSING

(defn process [year name file]
  (->> file
       (read-vacations name)
       flatten
       (create-events year)))

;; END PROCESSING

(defn -main [& args]
  (let [[year name file] args]
    (process year name file)))
