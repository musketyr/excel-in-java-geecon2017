package cz.orany.excel.geecon2017

import com.agorapulse.dru.Dru
import org.junit.ClassRule
import org.junit.Rule
import org.junit.rules.TemporaryFolder
import spock.lang.Shared
import spock.lang.Specification

import java.awt.*

class ExcelExporterSpec extends Specification {

    // https://agorapulse.github.io/dru/
    @Rule Dru dru = Dru.plan {
        from ('speakers.json') {
            map {
                to Speaker
            }
        }
    }

    @ClassRule @Shared TemporaryFolder tmp = new TemporaryFolder()

    ExcelExporter exporter = new ExcelExporter()

    void 'speakers loaded'() {
        expect:
            dru.findAllByType(Speaker).size() == 39
    }

    void 'basic export'() {
        given:
            File excel = tmp.newFile("${System.currentTimeMillis()}.xlsx")
        when:
            exporter.export(dru.findAllByType(Speaker), excel)
            open excel
        then:
            noExceptionThrown()
    }

    void cleanupSpec() {
        Thread.sleep(30000)
    }

    static void open(File file) {
        if (Desktop.desktopSupported && Desktop.desktop.isSupported(Desktop.Action.OPEN)) {
            Desktop.desktop.open(file)

        }
    }

}
