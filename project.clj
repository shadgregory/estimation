(defproject estimation "0.1.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :dependencies [[org.clojure/clojure "1.10.3"]
                 [org.apache.poi/poi "4.1.2"]
                 [org.apache.poi/poi-ooxml "4.1.2"]]
  :main ^:skip-aot estimation.core
  :target-path "target/%s"
  :profiles {:uberjar {:aot :all
                       :jvm-opts ["-Dclojure.compiler.direct-linking=true"]}})
