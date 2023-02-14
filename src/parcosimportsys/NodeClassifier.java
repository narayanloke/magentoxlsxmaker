/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package parcosimportsys;

/**
 *
 * @author narayan
 */
public class NodeClassifier {
    private int nodenumber;
    private String classifier;
    
    public NodeClassifier(int nodenumber, String classifier) {
        this.nodenumber = nodenumber;
        this.classifier = classifier;
    }
    

    public int getNodenumber() {
        return nodenumber;
    }

    public void setNodenumber(int nodenumber) {
        this.nodenumber = nodenumber;
    }

    public String getClassifier() {
        return classifier;
    }

    public void setClassifier(String classifier) {
        this.classifier = classifier;
    }
}
