package cz.orany.excel.geecon2017

import com.agorapulse.dru.Dru
import org.junit.ClassRule
import org.junit.Rule
import org.junit.rules.TemporaryFolder
import spock.lang.IgnoreRest
import spock.lang.Shared
import spock.lang.Specification

import java.awt.Desktop

class ExcelExporterSpec extends Specification {

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
            File excel = tmp.newFile("basic${System.currentTimeMillis()}.xlsx")
        when:
            exporter.basicExport(dru.findAllByType(Speaker), excel)
            open excel
        then:
            noExceptionThrown()
    }

    void 'export with filter'() {
        given:
            File excel = tmp.newFile("filtered${System.currentTimeMillis()}.xlsx")
        when:
            exporter.exportWithFilter(dru.findAllByType(Speaker), excel)
            open excel
        then:
            noExceptionThrown()
    }

    void 'export with filter and rowspan'() {
        given:
            File excel = tmp.newFile("rowspan${System.currentTimeMillis()}.xlsx")
        when:
            exporter.exportWithRowspan(dru.findAllByType(Speaker), excel)
            open excel
        then:
            noExceptionThrown()
    }

    void 'export with links and images'() {
        given:
            File excel = tmp.newFile("images${System.currentTimeMillis()}.xlsx")
        when:
            exporter.exportWithLinksAndImages(dru.findAllByType(Speaker), excel)
            open excel
        then:
            noExceptionThrown()
    }

    void 'export with comments'() {
        given:
            File excel = tmp.newFile("comments${System.currentTimeMillis()}.xlsx")
        when:
            exporter.exportWithComments(dru.findAllByType(Speaker), excel)
            open excel
        then:
            noExceptionThrown()
    }

    void 'export for print'() {
        given:
            File excel = tmp.newFile("print${System.currentTimeMillis()}.xlsx")
        when:
            exporter.exportPrint(dru.findAllByType(Speaker), excel)
            open excel
        then:
            noExceptionThrown()
    }

    void 'export for print with styles'() {
        given:
            File excel = tmp.newFile("print${System.currentTimeMillis()}.xlsx")
        when:
            exporter.exportPrintWithStyles(dru.findAllByType(Speaker), excel)
            open excel
        then:
            noExceptionThrown()
    }

    void 'load ysoft'() {
        given:
            File excel = tmp.newFile("print${System.currentTimeMillis()}.xlsx")
        when:
            exporter.exportPrintWithStyles(dru.findAllByType(Speaker), excel)
            Collection<Speaker> speakers = exporter.loadSpeakersData(excel.newInputStream())
        then:
            noExceptionThrown()
            speakers
            speakers.size() == 3
            speakers.first().name == 'Jakub Fojtl'
            speakers.first().presentations.first().title == 'Infinispan in the world dominated by raft'
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
