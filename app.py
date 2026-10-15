                except Exception as error:

                    st.error("Your answer could not be saved.")
                    st.exception(error)
                    st.stop()

                if is_correct:

                    st.session_state.current_correct = True
                    st.session_state.last_feedback = "✓ Correct!"
                    st.session_state.last_feedback_type = "correct"
                    st.session_state.show_explanation = True
                    st.session_state.flash_type = "correct"
                    st.session_state.wrong_answer_streak = 0
                    st.session_state.meltdown_popup = False

                else:

                    st.session_state.last_feedback = (
                        "✗ Incorrect — try again."
                    )

                    st.session_state.last_feedback_type = "incorrect"
                    st.session_state.flash_type = "incorrect"

                    st.session_state.wrong_answer_streak += 1

                    if st.session_state.wrong_answer_streak >= 2:
                        st.session_state.meltdown_popup = True

                st.rerun()

            # For "toggle_crossout", there is intentionally no
            # answer-submission logic here. The component has already
            # updated its visual state and Streamlit stores that state
            # under component_key.


    # ============================================================
    # FEEDBACK
    # ============================================================

    if st.session_state.last_feedback:

        if st.session_state.last_feedback_type == "correct":

            st.success(
                st.session_state.last_feedback
            )

        elif st.session_state.last_feedback_type == "incorrect":

            st.error(
                st.session_state.last_feedback
            )


    # ============================================================
    # EXPLANATION
    # ============================================================

    if st.session_state.show_explanation:

        st.divider()

        st.subheader("Explanation")

        st.info(
            question["explanation"]
        )

        st.write(
            f"**Correct answer: "
            f"{question['correct_letter']}. "
            f"{question['correct_answer']}**"
        )

        st.divider()

        is_last_in_queue = (
            current_index == len(questions) - 1
        )

        next_button_text = (
            "Finish Quiz"
            if is_last_in_queue
            else "Next Question →"
        )

        if st.button(
            next_button_text,
            type="primary",
            use_container_width=True
        ):

            if is_last_in_queue:
                st.session_state.quiz_complete = True
            else:
                st.session_state.current_index += 1

            st.session_state.current_correct = False
            st.session_state.last_feedback = ""
            st.session_state.last_feedback_type = ""
            st.session_state.show_explanation = False
            st.session_state.flash_type = None
            st.session_state.wrong_answer_streak = 0
            st.session_state.meltdown_popup = False

            st.rerun()
