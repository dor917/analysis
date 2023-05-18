package com.daon.analysis.dto;

public class DuplicationData {
    private String pointName;
    private String treeName;

    public String getPointName() {
        return pointName;
    }

    public void setPointName(String pointName) {
        this.pointName = pointName;
    }

    public String getTreeName() {
        return treeName;
    }

    public void setTreeName(String treeName) {
        this.treeName = treeName;
    }

    @Override
    public String toString() {
        return "DuplicationData{" +
                "pointName='" + pointName + '\'' +
                ", treeName='" + treeName + '\'' +
                '}';
    }

    @Override
    public int hashCode() {
        return String.valueOf(this.pointName).hashCode() + String.valueOf(this.treeName).hashCode();
    }

    @Override
    public boolean equals(Object obj) {
        if (obj instanceof DuplicationData) {
            DuplicationData temp = (DuplicationData) obj;
            return temp.pointName == this.pointName && temp.treeName == this.treeName;
        }
        return false;
    }
}
