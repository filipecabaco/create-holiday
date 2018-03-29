(defproject registration "1.0.0"
  :dependencies [[org.clojure/clojure "1.8.0"]
                 [dk.ative/docjure "1.12.0"]
                 [org.apache.poi/poi "3.17"]
                 [ring "1.6.3"]
                 [commons-io/commons-io "2.5"]]

  :main ^:skip-aot registration.core
  :target-path "target/%s"
  :profiles {:uberjar {:aot :all}})
