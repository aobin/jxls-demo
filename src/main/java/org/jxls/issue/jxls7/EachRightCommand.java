package org.jxls.issue.jxls7;

import org.jxls.command.EachCommand;

public class EachRightCommand extends EachCommand {

    public void setDir(String direction) {
        if ("RIGHT".equalsIgnoreCase(direction)) {
            this.setDirection(Direction.RIGHT);
        } else if("DOWN".equalsIgnoreCase(direction)) {
            this.setDirection(Direction.DOWN);
        }
    }

}
