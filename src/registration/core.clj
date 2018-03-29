(ns registration.core
  (:gen-class)
  (:require [dk.ative.docjure.spreadsheet :as ss]
            [clojure.java.io :as io]
            [clojure.string :as string])
  (:import [org.apache.commons.io IOUtils]
           [java.time LocalDate]))

(def months '("Jan" "Feb" "Mar" "Apr" "May" "Jun" "Jul" "Aug" "Sep" "Oct" "Nov" "Dec"))

;; BEGIN READ VACATION SHEET
(defn get-month-number [month]
  (case month
    "Jan" 1
    "Feb" 2
    "Mar" 3
    "Apr" 4
    "May" 5
    "Jun" 6
    "Jul" 7
    "Aug" 8
    "Sep" 9
    "Oct" 10
    "Nov" 11
    "Dec" 12))

(defn is-number [cell]
  (try
    (not (zero? (.getNumericCellValue cell)))
    (catch Exception e false)))

(defn not-number [cell]
  (not (is-number cell)))

(defn get-days [row]
  (let [all (reverse (iterator-seq (.cellIterator row)))
        remove-filler (drop-while not-number all)
        {days true empty false} (group-by is-number remove-filler)
        empty-count (count empty)
        days-count (count days)]
    {:empty empty-count :days days-count}))

(defn filter-by-name [name row]
  (let [cell (.getCell row 0)
        cellName (.getStringCellValue cell)]
    (= cellName name)))

(defn process-cell [cell]
  (let [value (.getStringCellValue cell)]
    (cond
      (re-matches #"(?i).*AL.*" value) :vacation
      (re-matches #"(?i).*SL.*" value) :sick
      (= "-" value) :holiday
      :else :full)))

(defn process-row [skip month-days month row]
  (let [all (iterator-seq (.cellIterator row))
        [name-cell] all
        name (.getStringCellValue name-cell)
        skipped (drop skip all)
        cells (take month-days skipped)
        days (if-not (empty? name) (seq (map process-cell cells)) nil)]
    (if-not (nil? days)
      {:name name :days days :month (get-month-number month)}
      nil)))

(defn read-sheet [name workbook month]
  (let [sheet (ss/select-sheet month workbook)
        rows  (ss/row-seq sheet)
        [month-row & rows] (drop 4 rows)
        {empty :empty days :days} (get-days month-row)
        selected (filter (partial filter-by-name name) rows)]
    (map (partial process-row empty days month) selected)))

(defn read-vacation [name file]
  (with-open [is (io/input-stream file)
              workbook (ss/load-workbook is)]
    (map (partial read-sheet name workbook) months)))

(defn read-vacations [name file]
  (read-vacation name file))

;; END READ VACATION SHEET

;; START CREATE EVENT
(defn is-full [all]
  (let [{status :status} all]
    (= status :full)))

(defn create-day [year month [day status]]
  {:year year
   month :month
   :day (inc day)
   :status status})

(defn create-event [year schedule]
  (let [{days :days month :month} schedule
        days-with-index (map-indexed vector days)
        status-map (map (partial create-day year month) days-with-index)
        filtered-status-map (filter (complement is-full) status-map)]
    (println filtered-status-map)))

(defn create-events [year schedule]
  (doall (map (partial create-event year) schedule)))
;; END CREATE EVENT

;; START PROCESSING

(defn process [year name vacation]
  (->> vacation
    (read-vacations name)
    flatten
    (create-events year)))

;; END PROCESSING

(defn cli-handler [year name vacation]
  (process year name vacation))

(defn -main [& args]
  (let [[year name vacation] args]
    (cli-handler year name vacation)))
